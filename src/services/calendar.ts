import type { AuthProvider } from "../types/auth.js";
import type {
  GraphEvent,
  GraphFindMeetingTimesResponse,
  GraphPagedResponse,
} from "../types/graph.js";
import { graphFetch } from "../utils/errors.js";
import {
  SCOPES,
  DEFAULT_PAGE_SIZE_LARGE,
  DEFAULT_TIMEZONE,
  DEFAULT_MEETING_DURATION_MINUTES,
} from "../constants.js";

export interface ListEventsParams {
  startDateTime: string;
  endDateTime: string;
  top?: number;
}

export interface CreateEventParams {
  subject: string;
  start: string;
  end: string;
  timeZone?: string;
  body?: string;
  location?: string;
  attendees?: string[];
  isOnlineMeeting?: boolean;
}

export interface UpdateEventParams {
  eventId: string;
  subject?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  body?: string;
  location?: string;
}

export interface FindFreeSlotsParams {
  attendees: string[];
  startDateTime: string;
  endDateTime: string;
  durationMinutes?: number;
}

export function createCalendarService(auth: AuthProvider) {
  async function getToken() {
    return auth.getAccessToken([...SCOPES.CALENDAR]);
  }

  async function listEvents(params: ListEventsParams): Promise<GraphEvent[]> {
    const token = await getToken();
    const { startDateTime, endDateTime, top = DEFAULT_PAGE_SIZE_LARGE } = params;

    const queryParams = new URLSearchParams({
      startDateTime,
      endDateTime,
      $top: String(top),
      $select:
        "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,bodyPreview",
      $orderby: "start/dateTime",
    });

    const result = await graphFetch<GraphPagedResponse<GraphEvent>>(
      token,
      `/me/calendarView?${queryParams}`
    );
    return result.value;
  }

  async function createEvent(params: CreateEventParams): Promise<GraphEvent> {
    const token = await getToken();
    const {
      subject,
      start,
      end,
      timeZone = DEFAULT_TIMEZONE,
      body,
      location,
      attendees,
      isOnlineMeeting = false,
    } = params;

    const eventBody: Record<string, unknown> = {
      subject,
      start: { dateTime: start, timeZone },
      end: { dateTime: end, timeZone },
      isOnlineMeeting,
    };

    if (body) {
      eventBody.body = { contentType: "Text", content: body };
    }
    if (location) {
      eventBody.location = { displayName: location };
    }
    if (attendees?.length) {
      eventBody.attendees = attendees.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      }));
    }

    return graphFetch<GraphEvent>(token, "/me/events", {
      method: "POST",
      body: JSON.stringify(eventBody),
    });
  }

  async function updateEvent(params: UpdateEventParams): Promise<GraphEvent> {
    const token = await getToken();
    const { eventId, timeZone = DEFAULT_TIMEZONE, ...updates } = params;

    const eventBody: Record<string, unknown> = {};
    if (updates.subject) eventBody.subject = updates.subject;
    if (updates.start) eventBody.start = { dateTime: updates.start, timeZone };
    if (updates.end) eventBody.end = { dateTime: updates.end, timeZone };
    if (updates.body) eventBody.body = { contentType: "Text", content: updates.body };
    if (updates.location) eventBody.location = { displayName: updates.location };

    return graphFetch<GraphEvent>(token, `/me/events/${eventId}`, {
      method: "PATCH",
      body: JSON.stringify(eventBody),
    });
  }

  async function deleteEvent(eventId: string): Promise<{ success: true }> {
    const token = await getToken();
    await graphFetch(token, `/me/events/${eventId}`, { method: "DELETE" });
    return { success: true };
  }

  async function findFreeSlots(
    params: FindFreeSlotsParams
  ): Promise<GraphFindMeetingTimesResponse> {
    const token = await getToken();
    const {
      attendees,
      startDateTime,
      endDateTime,
      durationMinutes = DEFAULT_MEETING_DURATION_MINUTES,
    } = params;

    return graphFetch<GraphFindMeetingTimesResponse>(token, "/me/findMeetingTimes", {
      method: "POST",
      body: JSON.stringify({
        attendees: attendees.map((email) => ({
          emailAddress: { address: email },
          type: "required",
        })),
        timeConstraint: {
          timeslots: [
            {
              start: { dateTime: startDateTime, timeZone: DEFAULT_TIMEZONE },
              end: { dateTime: endDateTime, timeZone: DEFAULT_TIMEZONE },
            },
          ],
        },
        meetingDuration: `PT${durationMinutes}M`,
      }),
    });
  }

  return { listEvents, createEvent, updateEvent, deleteEvent, findFreeSlots };
}

export type CalendarService = ReturnType<typeof createCalendarService>;
