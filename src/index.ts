#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import { OutlookCalendarClient, CalendarEvent } from './outlook/calendar-client.js';

// Check if running on Windows
if (process.platform !== 'win32') {
  console.error('Error: This MCP server only works on Windows');
  console.error('It uses COM Interop to communicate with Outlook, which is Windows-only');
  process.exit(1);
}

// Initialize Outlook client
const outlookClient = new OutlookCalendarClient();

// Define MCP tools
const TOOLS = [
  {
    name: 'outlook_list_events',
    description: 'List calendar events from Outlook. Optionally filter by date range (ISO 8601 format: YYYY-MM-DDTHH:mm:ss)',
    inputSchema: {
      type: 'object',
      properties: {
        startDate: {
          type: 'string',
          description: 'Start date/time in ISO 8601 format (e.g., "2024-01-01T00:00:00")'
        },
        endDate: {
          type: 'string',
          description: 'End date/time in ISO 8601 format (e.g., "2024-12-31T23:59:59")'
        }
      }
    }
  },
  {
    name: 'outlook_get_event',
    description: 'Get details of a specific calendar event by ID',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: {
          type: 'string',
          description: 'The EntryID of the event to retrieve'
        }
      },
      required: ['eventId']
    }
  },
  {
    name: 'outlook_create_event',
    description: 'Create a new calendar event in Outlook',
    inputSchema: {
      type: 'object',
      properties: {
        subject: {
          type: 'string',
          description: 'The title/subject of the event'
        },
        start: {
          type: 'string',
          description: 'Start date/time in ISO 8601 format'
        },
        end: {
          type: 'string',
          description: 'End date/time in ISO 8601 format'
        },
        body: {
          type: 'string',
          description: 'Event description/body'
        },
        location: {
          type: 'string',
          description: 'Event location'
        },
        attendees: {
          type: 'array',
          description: 'List of attendee email addresses',
          items: {
            type: 'string'
          }
        },
        isAllDay: {
          type: 'boolean',
          description: 'Whether this is an all-day event',
          default: false
        }
      },
      required: ['subject', 'start', 'end']
    }
  },
  {
    name: 'outlook_update_event',
    description: 'Update an existing calendar event',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: {
          type: 'string',
          description: 'The EntryID of the event to update'
        },
        subject: {
          type: 'string',
          description: 'The title/subject of the event'
        },
        start: {
          type: 'string',
          description: 'Start date/time in ISO 8601 format'
        },
        end: {
          type: 'string',
          description: 'End date/time in ISO 8601 format'
        },
        body: {
          type: 'string',
          description: 'Event description/body'
        },
        location: {
          type: 'string',
          description: 'Event location'
        }
      },
      required: ['eventId']
    }
  },
  {
    name: 'outlook_delete_event',
    description: 'Delete a calendar event from Outlook',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: {
          type: 'string',
          description: 'The EntryID of the event to delete'
        }
      },
      required: ['eventId']
    }
  },
  {
    name: 'outlook_search_events',
    description: 'Search calendar events by query string (searches in subject and body)',
    inputSchema: {
      type: 'object',
      properties: {
        query: {
          type: 'string',
          description: 'Search query to find events'
        }
      },
      required: ['query']
    }
  }
];

// Initialize MCP server
const server = new Server(
  {
    name: 'mcp-server-outlook',
    version: '1.0.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// List available tools
server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: TOOLS,
  };
});

// Handle tool calls
server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const { name, arguments: args } = request.params;
    let result: any;

    switch (name) {
      case 'outlook_list_events':
        result = await outlookClient.listEvents(args.startDate, args.endDate);
        break;

      case 'outlook_get_event':
        result = await outlookClient.getEvent(args.eventId);
        break;

      case 'outlook_create_event':
        const newEvent: CalendarEvent = {
          subject: args.subject,
          start: args.start,
          end: args.end,
          body: args.body,
          location: args.location,
          attendees: args.attendees,
          isAllDay: args.isAllDay || false
        };
        result = await outlookClient.createEvent(newEvent);
        break;

      case 'outlook_update_event':
        const updates: Partial<CalendarEvent> = {};
        if (args.subject) updates.subject = args.subject;
        if (args.start) updates.start = args.start;
        if (args.end) updates.end = args.end;
        if (args.body) updates.body = args.body;
        if (args.location) updates.location = args.location;
        result = await outlookClient.updateEvent(args.eventId, updates);
        break;

      case 'outlook_delete_event':
        result = await outlookClient.deleteEvent(args.eventId);
        break;

      case 'outlook_search_events':
        result = await outlookClient.searchEvents(args.query);
        break;

      default:
        throw new Error(`Unknown tool: ${name}`);
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify(result, null, 2),
        },
      ],
    };
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${errorMessage}`,
        },
      ],
      isError: true,
    };
  }
});

// Start server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('MCP Outlook Calendar Server running on stdio');
  console.error('Platform:', process.platform);
  console.error('Note: Outlook must be installed and accessible on this system');
}

main().catch((error) => {
  console.error('Fatal error:', error);
  process.exit(1);
});
