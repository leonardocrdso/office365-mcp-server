import type { GraphEvent, GraphFindMeetingTimesResponse } from "../types/graph.js";
import { recipientAddress } from "../types/graph.js";

export function formatEventList(events: GraphEvent[]): string {
  if (events.length === 0) return "Nenhum evento encontrado no período.";

  const formatted = events.map((e) => {
    const start = new Date(e.start.dateTime).toLocaleString("pt-BR");
    const end = new Date(e.end.dateTime).toLocaleString("pt-BR");
    const location = e.location?.displayName ? ` @ ${e.location.displayName}` : "";
    const online = e.isOnlineMeeting ? " [Online]" : "";
    const attendees = e.attendees?.length
      ? `\n  Participantes: ${e.attendees.map((a) => recipientAddress(a)).join(", ")}`
      : "";

    return `- **${e.subject}**${online}${location}\n  ${start} → ${end}${attendees}\n  ID: ${e.id}`;
  });

  return `## Eventos (${formatted.length})\n\n${formatted.join("\n\n")}`;
}

export function formatCreatedEvent(event: GraphEvent): string {
  let text = `Evento criado com sucesso!\n\n**${event.subject}**\nID: ${event.id}`;
  if (event.onlineMeeting?.joinUrl) {
    text += `\nLink da reunião: ${event.onlineMeeting.joinUrl}`;
  }
  return text;
}

export function formatFreeSlotsResult(result: GraphFindMeetingTimesResponse): string {
  const suggestions = result.meetingTimeSuggestions ?? [];

  if (suggestions.length === 0) {
    return "Nenhum horário disponível encontrado para todos os participantes no período.";
  }

  const formatted = suggestions.map((s, i) => {
    const start = new Date(s.meetingTimeSlot.start.dateTime).toLocaleString("pt-BR");
    const end = new Date(s.meetingTimeSlot.end.dateTime).toLocaleString("pt-BR");
    return `${i + 1}. ${start} → ${end} (confiança: ${s.confidence}%)`;
  });

  return `## Horários Sugeridos\n\n${formatted.join("\n")}`;
}
