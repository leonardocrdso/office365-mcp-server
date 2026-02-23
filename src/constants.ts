export const DEFAULT_PAGE_SIZE_SMALL = 10;
export const DEFAULT_PAGE_SIZE_LARGE = 20;
export const DEFAULT_MEETING_DURATION_MINUTES = 30;
export const DEFAULT_TIMEZONE = "America/Sao_Paulo";
export const BODY_PREVIEW_MAX_LENGTH = 100;
export const MESSAGE_CONTENT_MAX_LENGTH = 200;

export const SCOPES = {
  MAIL: ["Mail.Read", "Mail.Send", "Mail.ReadWrite"],
  CALENDAR: ["Calendars.Read", "Calendars.ReadWrite"],
  DRIVE: ["Files.Read.All", "Files.ReadWrite.All"],
  SHAREPOINT: ["Sites.Read.All", "Sites.ReadWrite.All"],
  TEAMS: [
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Send",
    "Chat.Read",
    "Chat.ReadWrite",
  ],
} as const;
