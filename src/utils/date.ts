import { DEFAULT_TIMEZONE } from "../constants.js";

export function formatDateBR(isoDate: string): string {
  return new Date(isoDate).toLocaleString("pt-BR", { timeZone: DEFAULT_TIMEZONE });
}

/**
 * Formats a datetime string returned by Graph API when using the
 * Prefer: outlook.timezone header. These strings have no timezone
 * suffix (e.g. "2025-01-20T10:00:00.0000000") and are already in
 * the preferred timezone, so we parse them directly to avoid
 * misinterpretation by the Date constructor.
 */
export function formatCalendarDate(dateTime: string): string {
  const match = dateTime.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  if (match) {
    const [, year, month, day, hour, minute] = match;
    return `${day}/${month}/${year}, ${hour}:${minute}`;
  }
  return new Date(dateTime).toLocaleString("pt-BR", { timeZone: DEFAULT_TIMEZONE });
}
