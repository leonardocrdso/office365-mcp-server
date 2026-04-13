import { readFile, writeFile } from "node:fs/promises";
import { resolveStoragePaths, ensureStorageDir } from "./storage.js";

export async function loadConfig(): Promise<{ clientId?: string; tenantId?: string }> {
  const { configPath } = resolveStoragePaths();
  try {
    return JSON.parse(await readFile(configPath, "utf-8"));
  } catch {
    return {};
  }
}

export async function saveConfig(clientId: string, tenantId: string): Promise<string> {
  await ensureStorageDir();
  const { configPath } = resolveStoragePaths();
  await writeFile(configPath, JSON.stringify({ clientId, tenantId }, null, 2), { mode: 0o600 });
  return configPath;
}
