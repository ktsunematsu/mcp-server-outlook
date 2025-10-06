import { PowerShellBridge } from './powershell-bridge.js';

export interface CalendarEvent {
  id?: string;
  subject: string;
  start: string;
  end: string;
  body?: string;
  location?: string;
  attendees?: string[];
  isAllDay?: boolean;
}

export class OutlookCalendarClient {
  private bridge: PowerShellBridge;

  constructor() {
    this.bridge = new PowerShellBridge();
  }

  async listEvents(startDate?: string, endDate?: string): Promise<CalendarEvent[]> {
    const params: Record<string, any> = {};
    if (startDate) params.StartDate = startDate;
    if (endDate) params.EndDate = endDate;

    const result = await this.bridge.executeScript('list', params);

    if (!result.success) {
      throw new Error(result.error || 'Failed to list events');
    }

    return result.data as CalendarEvent[];
  }

  async getEvent(eventId: string): Promise<CalendarEvent> {
    const result = await this.bridge.executeScript('get', { EventId: eventId });

    if (!result.success) {
      throw new Error(result.error || 'Failed to get event');
    }

    return result.data as CalendarEvent;
  }

  async createEvent(event: CalendarEvent): Promise<CalendarEvent> {
    const params: Record<string, any> = {
      Subject: event.subject,
      StartDate: event.start,
      EndDate: event.end
    };

    if (event.body) params.Body = event.body;
    if (event.location) params.Location = event.location;
    if (event.isAllDay) params.IsAllDay = event.isAllDay;
    if (event.attendees && event.attendees.length > 0) {
      params.Attendees = event.attendees.join(';');
    }

    const result = await this.bridge.executeScript('create', params);

    if (!result.success) {
      throw new Error(result.error || 'Failed to create event');
    }

    return result.data as CalendarEvent;
  }

  async updateEvent(eventId: string, updates: Partial<CalendarEvent>): Promise<CalendarEvent> {
    const params: Record<string, any> = { EventId: eventId };

    if (updates.subject) params.Subject = updates.subject;
    if (updates.start) params.StartDate = updates.start;
    if (updates.end) params.EndDate = updates.end;
    if (updates.body) params.Body = updates.body;
    if (updates.location) params.Location = updates.location;

    const result = await this.bridge.executeScript('update', params);

    if (!result.success) {
      throw new Error(result.error || 'Failed to update event');
    }

    return result.data as CalendarEvent;
  }

  async deleteEvent(eventId: string): Promise<{ success: boolean; message: string }> {
    const result = await this.bridge.executeScript('delete', { EventId: eventId });

    if (!result.success) {
      throw new Error(result.error || 'Failed to delete event');
    }

    return result.data as { success: boolean; message: string };
  }

  async searchEvents(query: string): Promise<CalendarEvent[]> {
    const result = await this.bridge.executeScript('search', { Query: query });

    if (!result.success) {
      throw new Error(result.error || 'Failed to search events');
    }

    return result.data as CalendarEvent[];
  }
}
