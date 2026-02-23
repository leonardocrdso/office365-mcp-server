import { msalClient } from "../auth/msal-client.js";
import { graphFetch } from "../utils/errors.js";

const CALENDAR_SCOPES = ["Calendars.Read", "Calendars.ReadWrite"];

async function getToken() {
  return msalClient.getAccessToken(CALENDAR_SCOPES);
}

export interface ListEventsParams {
  startDateTime: string;
  endDateTime: string;
  top?: number;
}

export async function listEvents(params: ListEventsParams) {
  const token = await getToken();
  const { startDateTime, endDateTime, top = 20 } = params;

  const queryParams = new URLSearchParams({
    startDateTime,
    endDateTime,
    $top: String(top),
    $select: "id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeetingUrl,bodyPreview",
    $orderby: "start/dateTime",
  });

  const result = await graphFetch(token, `/me/calendarView?${queryParams}`);
  return result.value;
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

export async function createEvent(params: CreateEventParams) {
  const token = await getToken();
  const {
    subject,
    start,
    end,
    timeZone = "America/Sao_Paulo",
    body,
    location,
    attendees,
    isOnlineMeeting = false,
  } = params;

  const eventBody: Record<string, any> = {
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

  const result = await graphFetch(token, "/me/events", {
    method: "POST",
    body: JSON.stringify(eventBody),
  });
  return result;
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

export async function updateEvent(params: UpdateEventParams) {
  const token = await getToken();
  const { eventId, timeZone = "America/Sao_Paulo", ...updates } = params;

  const eventBody: Record<string, any> = {};
  if (updates.subject) eventBody.subject = updates.subject;
  if (updates.start) eventBody.start = { dateTime: updates.start, timeZone };
  if (updates.end) eventBody.end = { dateTime: updates.end, timeZone };
  if (updates.body) eventBody.body = { contentType: "Text", content: updates.body };
  if (updates.location) eventBody.location = { displayName: updates.location };

  const result = await graphFetch(token, `/me/events/${eventId}`, {
    method: "PATCH",
    body: JSON.stringify(eventBody),
  });
  return result;
}

export async function deleteEvent(eventId: string) {
  const token = await getToken();
  await graphFetch(token, `/me/events/${eventId}`, { method: "DELETE" });
  return { success: true };
}

export interface FindFreeSlotsParams {
  attendees: string[];
  startDateTime: string;
  endDateTime: string;
  durationMinutes?: number;
}

export async function findFreeSlots(params: FindFreeSlotsParams) {
  const token = await getToken();
  const { attendees, startDateTime, endDateTime, durationMinutes = 30 } = params;

  const result = await graphFetch(token, "/me/findMeetingTimes", {
    method: "POST",
    body: JSON.stringify({
      attendees: attendees.map((email) => ({
        emailAddress: { address: email },
        type: "required",
      })),
      timeConstraint: {
        timeslots: [
          {
            start: { dateTime: startDateTime, timeZone: "America/Sao_Paulo" },
            end: { dateTime: endDateTime, timeZone: "America/Sao_Paulo" },
          },
        ],
      },
      meetingDuration: `PT${durationMinutes}M`,
    }),
  });
  return result;
}
