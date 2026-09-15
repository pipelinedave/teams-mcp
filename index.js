#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema
} from '@modelcontextprotocol/sdk/types.js';
import { browserManager } from './src/browserManager.js';
import { teamsClient } from './src/teamsClient.js';
import { config } from './src/config.js';

const server = new Server(
  {
    name: 'teams-mcp',
    version: '1.0.0'
  },
  {
    capabilities: {
      tools: {}
    }
  }
);

// Tenant-Parameter org-neutral bauen. Sind über TEAMS_MCP_TENANTS / TEAMS_MCP_TENANT_REALMS
// konfigurierte Tenants vorhanden, werden sie als validierte Enum angeboten; sonst ist
// es ein freies Textfeld (möglichst mit Realm, z.B. "meine-org.onmicrosoft.com").
const knownTenants = config.tenants;
const defaultTenant = config.defaultTenant || 'adesso';

const tenantParam = (options = {}) => {
  const base = {
    type: 'string',
    description: options.includeAll
      ? [knownTenants.length ? `Bekannte Tenants: ${knownTenants.join(', ')}. ` : '', `Wert oder "all" (Standard: "${options.def || (knownTenants.length ? 'all' : defaultTenant)}").`].join('')
      : [knownTenants.length ? `Bekannte Tenants: ${knownTenants.join(', ')}. ` : '', `Tenant (Name oder Realm, z.B. "meine-org.onmicrosoft.com", Standard: "${defaultTenant}").`].join('')
  };
  if (knownTenants.length) base.enum = [...knownTenants, ...(options.includeAll ? ['all'] : [])];
  if (options.def !== undefined) base.default = options.def;
  else base.default = defaultTenant;
  return base;
};

// Hilfsfunktion: Löst einen Tenant-Parameter smart auf.
// Kein Tenant oder "all" (bei Einzeltenant-Operationen) fällt sauber auf den Standard-Tenant zurück ("adesso").
function resolveTenant(tenant) {
  if (!tenant || tenant === 'all') {
    return defaultTenant;
  }
  return browserManager.normalizeTenant(tenant);
}

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'teams_status',
        description: 'Prüft den Anmeldestatus von Microsoft Teams für einen oder alle Tenants.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam({ includeAll: true, def: 'all' })
          }
        }
      },
      {
        name: 'teams_login',
        description: 'Öffnet ein sichtbares Browserfenster (WSLg) für die einmalige Anmeldung an einem Teams-Tenant.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam()
          },
          required: ['tenant']
        }
      },
      {
        name: 'teams_list_chats',
        description: 'Listet die neuesten 1:1- und Gruppen-Chats in Microsoft Teams mit Namen und Vorschau auf.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam(),
            limit: {
              type: 'number',
              description: 'Maximale Anzahl an Chats (Standard: 15)',
              default: 15
            }
          }
        }
      },
      {
        name: 'teams_get_messages',
        description: 'Liest die Nachrichtenverläufe eines Chats als Markdown aus (inkl. Autoren, Zeitstempel, Inhalt).',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam(),
            chat_index: {
              type: 'number',
              description: 'Index des Chats aus teams_list_chats (z.B. 0 für den obersten)'
            },
            chat_name: {
              type: 'string',
              description: 'Name des Chats oder Kollegen (z.B. "Marius", "Alexander Schnieders")'
            },
            limit: {
              type: 'number',
              description: 'Maximale Anzahl an Nachrichten (Standard: 20)',
              default: 20
            }
          }
        }
      },
      {
        name: 'teams_search',
        description: 'Durchsucht Teams-Nachrichten nach Stichwörtern (z.B. Kundenname, Ticketnummer, Kollege).',
        inputSchema: {
          type: 'object',
          properties: {
            query: {
              type: 'string',
              description: 'Suchbegriff (z.B. "UKO", "Hildburghausen", "00890612", "Gabriel")'
            },
            tenant: tenantParam(),
            limit: {
              type: 'number',
              description: 'Maximale Anzahl an Treffern (Standard: 10)',
              default: 10
            }
          },
          required: ['query']
        }
      },
      {
        name: 'teams_list_teams',
        description: 'Listet alle beigetretenen Microsoft Teams und deren Kanäle (Channels) auf.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam()
          }
        }
      },
      {
        name: 'teams_send_message',
        description: 'Sendet eine Nachricht in einen Microsoft Teams Chat (optional inklusive Dateianhängen).',
        inputSchema: {
          type: 'object',
          properties: {
            message: {
              type: 'string',
              description: 'Zu sendender Nachrichtentext'
            },
            chat_name: {
              type: 'string',
              description: 'Name des Empfänger-Chats oder Kollegen'
            },
            attachments: {
              type: 'array',
              items: { type: 'string' },
              description: 'Optionale Liste lokaler Dateipfade, die als Anhang mitgesendet werden sollen (z.B. ["/tmp/datei.zip"]).'
            },
            attachment: {
              type: 'string',
              description: 'Optionaler einzelner Dateipfad als Anhang (Alternative zu attachments).'
            },
            tenant: tenantParam()
          },
          required: ['message']
        }
      },
      {
        name: 'teams_inspect',
        description: 'Gibt technische Diagnose-Informationen über erkannte Teams-UI-Elemente, Chat-Pane, Compose-Footer, Buttons und Datei-Inputs zurück.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam()
          }
        }
      },
      {
        name: 'teams_close',
        description: 'Schließt die Hintergrund-Browserinstanzen von Microsoft Teams.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam({ includeAll: true, def: 'all' })
          }
        }
      },
      {
        name: 'teams_meeting_status',
        description: 'Ermittelt den aktuellen Meeting- und Anrufstatus (inkl. Meeting-Titel, Teilnehmer und aktive Sprecher).',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam()
          }
        }
      },
      {
        name: 'teams_start_tracking',
        description: 'Startet das DOM-basierte Active Speaker Tracking für das laufende Teams-Meeting.',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam(),
            output_path: {
              type: 'string',
              description: 'Optionaler Speicherpfad für die *.speakers.json (z.B. /pfad/meeting.speakers.json)'
            }
          }
        }
      },
      {
        name: 'teams_stop_tracking',
        description: 'Beendet das Active Speaker Tracking und exportiert die Sprecher-Timeline (*.speakers.json).',
        inputSchema: {
          type: 'object',
          properties: {
            tenant: tenantParam()
          }
        }
      }
    ]
  };
});

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      case 'teams_status': {
        const tenantArg = args?.tenant;
        if (!tenantArg || tenantArg === 'all') {
          // Falls Tenants konfiguriert sind, alle prüfen.
          // Falls keine Env-Konfiguration vorliegt, Smart-Default auf den Standard-Tenant ("adesso").
          const targets = knownTenants.length > 0 ? knownTenants : [defaultTenant];
          const statuses = {};
          for (const t of targets) {
            statuses[t] = await teamsClient.checkStatus(t);
          }
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify(statuses, null, 2)
              }
            ]
          };
        } else {
          const status = await teamsClient.checkStatus(resolveTenant(tenantArg));
          return {
            content: [{ type: 'text', text: JSON.stringify(status, null, 2) }]
          };
        }
      }

      case 'teams_login': {
        const tenant = resolveTenant(args?.tenant);
        await browserManager.openLoginWindow(tenant);
        const realm = browserManager.realm(tenant);
        return {
          content: [
            {
              type: 'text',
              text: `Ein sichtbares Browserfenster für Microsoft Teams (${tenant} - ${realm}) wurde geöffnet. Bitte melde dich dort einmalig an (inkl. MFA). Die Session wird im Profil dauerhaft gespeichert. Sobald du deine Chats siehst, ist teams-mcp für ${tenant} einsatzbereit.`
            }
          ]
        };
      }

      case 'teams_list_chats': {
        const tenant = resolveTenant(args?.tenant);
        const limit = args?.limit || 15;
        const result = await teamsClient.listChats(tenant, limit);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_get_messages': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.getMessages(tenant, {
          chatIndex: args?.chat_index,
          chatName: args?.chat_name,
          limit: args?.limit || 20
        });
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_search': {
        const tenant = resolveTenant(args?.tenant);
        const limit = args?.limit || 10;
        const result = await teamsClient.search(tenant, args.query, limit);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_list_teams': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.listTeams(tenant);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_send_message': {
        const tenant = resolveTenant(args?.tenant);
        const rawAttachments = args?.attachments || (args?.attachment ? [args.attachment] : []);
        const result = await teamsClient.sendMessage(tenant, {
          message: args.message,
          chatName: args.chat_name,
          attachments: rawAttachments
        });
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_inspect': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.inspectCompose(tenant);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_meeting_status': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.getMeetingStatus(tenant);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_start_tracking': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.startSpeakerTracking(tenant, args?.output_path);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_stop_tracking': {
        const tenant = resolveTenant(args?.tenant);
        const result = await teamsClient.stopSpeakerTracking(tenant);
        return {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }]
        };
      }

      case 'teams_close': {
        const tenant = args?.tenant || 'all';
        await browserManager.close(tenant);
        return {
          content: [
            {
              type: 'text',
              text: `Teams Browserinstanz (${tenant}) wurde geschlossen.`
            }
          ]
        };
      }

      default:
        throw new Error(`Unbekanntes Tool: ${name}`);
    }
  } catch (error) {
    return {
      isError: true,
      content: [{ type: 'text', text: `Fehler in ${name}: ${error.message}` }]
    };
  }
});

async function run() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

run().catch((err) => {
  console.error('Fatal error in teams-mcp:', err);
  process.exit(1);
});
