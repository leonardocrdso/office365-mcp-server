export const DEFAULT_PAGE_SIZE_SMALL = 10;
export const DEFAULT_PAGE_SIZE_LARGE = 20;
export const DEFAULT_MEETING_DURATION_MINUTES = 30;
export const DEFAULT_TIMEZONE = "America/Sao_Paulo";
export const PREFER_TIMEZONE_HEADER = `outlook.timezone="${DEFAULT_TIMEZONE}"`;
export const BODY_PREVIEW_MAX_LENGTH = 100;
export const MESSAGE_CONTENT_MAX_LENGTH = 200;
export const DEFAULT_EMAIL_BODY_MAX_LENGTH = 4000;
export const MAX_EMAIL_BODY_MAX_LENGTH = 12000;

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

export const GRAPH_ATTEMPT_TIMEOUT_MS = 40_000;
export const GRAPH_TOTAL_BUDGET_MS = 60_000;
export const GRAPH_MIN_ATTEMPT_BUDGET_MS = 5_000;
export const GRAPH_MAX_ATTEMPTS = 3;
export const GRAPH_RETRY_BASE_DELAY_MS = 1_000;
export const GRAPH_RETRY_MAX_DELAY_MS = 8_000;

export const DOWNLOAD_MAX_BYTES = 100 * 1024 * 1024;
export const DOWNLOAD_RETENTION_MS = 60 * 60 * 1000;
export const DOWNLOAD_SWEEP_INTERVAL_MS = 15 * 60 * 1000;
