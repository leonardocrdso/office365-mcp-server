import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import type { CalendarService } from "../services/calendar.js";
import { safeTool } from "../utils/errors.js";
import {
  formatEventList,
  formatCreatedEvent,
  formatFreeSlotsResult,
} from "../formatters/calendar.js";

export function registerCalendarTools(server: McpServer, calendar: CalendarService) {
  server.tool(
    "list-events",
    "Lista eventos do calendário em um intervalo de datas.",
    {
      startDateTime: z.string().describe("Data/hora início em ISO 8601 (ex: 2025-01-20T00:00:00)"),
      endDateTime: z.string().describe("Data/hora fim em ISO 8601 (ex: 2025-01-27T23:59:59)"),
      top: z.number().optional().describe("Número máximo de eventos (padrão: 20)"),
    },
    safeTool(async (params) => {
      const events = await calendar.listEvents(params);
      return {
        content: [{ type: "text" as const, text: formatEventList(events) }],
      };
    })
  );

  server.tool(
    "create-event",
    "Cria um novo evento/reunião no calendário.",
    {
      subject: z.string().describe("Título do evento"),
      start: z.string().describe("Data/hora início (ex: 2025-01-20T10:00:00)"),
      end: z.string().describe("Data/hora fim (ex: 2025-01-20T11:00:00)"),
      timeZone: z.string().optional().describe("Fuso horário (padrão: America/Sao_Paulo)"),
      body: z.string().optional().describe("Descrição do evento"),
      location: z.string().optional().describe("Local do evento"),
      attendees: z.array(z.string()).optional().describe("Emails dos participantes"),
      isOnlineMeeting: z.boolean().optional().describe("Se true, cria reunião online (Teams)"),
    },
    safeTool(async (params) => {
      const event = await calendar.createEvent(params);
      return {
        content: [{ type: "text" as const, text: formatCreatedEvent(event) }],
      };
    })
  );

  server.tool(
    "update-event",
    "Atualiza um evento existente no calendário.",
    {
      eventId: z.string().describe("ID do evento a atualizar"),
      subject: z.string().optional().describe("Novo título"),
      start: z.string().optional().describe("Nova data/hora início"),
      end: z.string().optional().describe("Nova data/hora fim"),
      timeZone: z.string().optional().describe("Fuso horário"),
      body: z.string().optional().describe("Nova descrição"),
      location: z.string().optional().describe("Novo local"),
    },
    safeTool(async (params) => {
      const event = await calendar.updateEvent(params);
      return {
        content: [
          {
            type: "text" as const,
            text: `Evento atualizado: **${event.subject}**`,
          },
        ],
      };
    })
  );

  server.tool(
    "delete-event",
    "Remove um evento do calendário.",
    {
      eventId: z.string().describe("ID do evento a remover"),
    },
    safeTool(async (params) => {
      await calendar.deleteEvent(params.eventId);
      return {
        content: [
          { type: "text" as const, text: "Evento removido com sucesso." },
        ],
      };
    })
  );

  server.tool(
    "find-free-slots",
    "Verifica disponibilidade de horários para reunião com participantes.",
    {
      attendees: z.array(z.string()).describe("Emails dos participantes"),
      startDateTime: z.string().describe("Início do período de busca (ISO 8601)"),
      endDateTime: z.string().describe("Fim do período de busca (ISO 8601)"),
      durationMinutes: z.number().optional().describe("Duração da reunião em minutos (padrão: 30)"),
    },
    safeTool(async (params) => {
      const result = await calendar.findFreeSlots(params);
      return {
        content: [{ type: "text" as const, text: formatFreeSlotsResult(result) }],
      };
    })
  );
}
