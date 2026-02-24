import { DEFAULT_TIMEZONE } from "../constants.js";

export function formatDateBR(isoDate: string): string {
  return new Date(isoDate).toLocaleString("pt-BR", { timeZone: DEFAULT_TIMEZONE });
}
