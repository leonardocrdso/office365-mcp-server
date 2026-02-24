import { formatGraphError } from "./errors.js";

type ToolResult = { content: { type: "text"; text: string }[]; isError?: boolean };
type ToolHandler<T extends unknown[] = unknown[]> = (...args: T) => Promise<ToolResult>;

export function safeTool<T extends unknown[]>(handler: ToolHandler<T>): ToolHandler<T> {
  return (async (...args: T) => {
    try {
      return await handler(...args);
    } catch (error) {
      const message = formatGraphError(error);
      return {
        content: [{ type: "text" as const, text: message }],
        isError: true,
      };
    }
  }) as ToolHandler<T>;
}
