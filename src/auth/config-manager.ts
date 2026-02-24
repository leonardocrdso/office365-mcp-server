import { readFile, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";

export const CONFIG_PATH = join(homedir(), ".office365-mcp-config.json");

export async function loadConfig(): Promise<{ clientId?: string; tenantId?: string }> {
  try {
    const data = await readFile(CONFIG_PATH, "utf-8");
    return JSON.parse(data);
  } catch {
    return {};
  }
}

export async function saveConfig(clientId: string, tenantId: string): Promise<void> {
  await writeFile(
    CONFIG_PATH,
    JSON.stringify({ clientId, tenantId }, null, 2),
    { mode: 0o600 }
  );
}
