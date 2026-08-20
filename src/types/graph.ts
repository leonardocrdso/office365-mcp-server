// Mail
export interface GraphEmailAddress {
  address: string;
  name?: string;
}

export interface GraphRecipient {
  emailAddress: GraphEmailAddress;
}

export interface GraphAttachment {
  name: string;
  contentType: string;
  size: number;
}

export interface GraphEmailMessage {
  id: string;
  conversationId?: string;
  subject: string;
  from: GraphRecipient;
  toRecipients: GraphRecipient[];
  ccRecipients?: GraphRecipient[];
  receivedDateTime: string;
  isRead: boolean;
  hasAttachments: boolean;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  attachments?: GraphAttachment[];
}

export interface GraphMailFolder {
  id: string;
  displayName: string;
  totalItemCount: number;
  unreadItemCount: number;
}

// Calendar
export interface GraphDateTimeZone {
  dateTime: string;
  timeZone: string;
}

export interface GraphEvent {
  id: string;
  subject: string;
  start: GraphDateTimeZone;
  end: GraphDateTimeZone;
  location?: { displayName?: string };
  organizer?: GraphRecipient;
  attendees?: Array<{ emailAddress: GraphEmailAddress; type: string }>;
  isOnlineMeeting?: boolean;
  onlineMeetingUrl?: string;
  onlineMeeting?: { joinUrl?: string };
  bodyPreview?: string;
  body?: { contentType: string; content: string };
}

export interface GraphMeetingTimeSuggestion {
  meetingTimeSlot: { start: GraphDateTimeZone; end: GraphDateTimeZone };
  confidence: number;
}

export interface GraphFindMeetingTimesResponse {
  meetingTimeSuggestions: GraphMeetingTimeSuggestion[];
}

// OneDrive
export interface GraphDriveItem {
  id: string;
  name: string;
  size?: number;
  lastModifiedDateTime: string;
  folder?: { childCount: number };
  file?: { mimeType: string };
  webUrl: string;
  parentReference?: { path?: string; driveId?: string; siteId?: string };
}

export interface GraphSharingLink {
  link: { webUrl: string; type: string; scope: string };
}

// SharePoint
export interface GraphSite {
  id: string;
  displayName: string;
  name: string;
  webUrl: string;
  description?: string;
}

export interface GraphDrive {
  id: string;
  name: string;
  driveType: string;
  webUrl: string;
  description?: string;
}

export interface GraphSearchHitResource {
  "@odata.type"?: string;
  id?: string;
  name?: string;
  displayName?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { displayName?: string } };
  parentReference?: {
    driveId?: string;
    siteId?: string;
  };
  size?: number;
}

export interface GraphSearchHit {
  hitId: string;
  rank?: number;
  summary?: string;
  resource: GraphSearchHitResource;
}

export interface GraphSearchHitsContainer {
  hits: GraphSearchHit[];
}

export interface GraphSearchResponse {
  hitsContainers: GraphSearchHitsContainer[];
}

// Teams
export interface GraphTeam {
  id: string;
  displayName: string;
  description?: string;
}

export interface GraphChannel {
  id: string;
  displayName: string;
  description?: string;
  membershipType: string;
}

export interface GraphChatMessage {
  id: string;
  createdDateTime: string;
  from?: { user?: { displayName?: string } };
  body?: { contentType: string; content: string };
}

export interface GraphChat {
  id: string;
  topic?: string;
  chatType: string;
  lastUpdatedDateTime?: string;
  members?: Array<{ displayName: string; email?: string }>;
}

// Generic
export interface GraphPagedResponse<T> {
  value: T[];
}

// Helpers
export function recipientAddress(recipient?: GraphRecipient): string {
  return recipient?.emailAddress?.address ?? "unknown";
}

export function recipientAddresses(recipients?: GraphRecipient[]): string {
  return recipients?.map((r) => r.emailAddress?.address).join(", ") ?? "";
}

export function messageSenderName(message: GraphChatMessage): string {
  return message.from?.user?.displayName ?? "Desconhecido";
}

export function toRecipient(email: string): GraphRecipient {
  return { emailAddress: { address: email } };
}

export function toAttendee(email: string, type: string = "required") {
  return { emailAddress: { address: email } as GraphEmailAddress, type };
}
