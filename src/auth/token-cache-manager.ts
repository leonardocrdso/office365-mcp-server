import { readFile, writeFile, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";

export const TOKEN_CACHE_PATH = join(homedir(), ".office365-mcp-tokens.json");

export async function loadTokenCache(): Promise<string | null> {
  try {
    return await readFile(TOKEN_CACHE_PATH, "utf-8");
  } catch {
    return null;
  }
}

export async function saveTokenCache(data: string): Promise<void> {
  await writeFile(TOKEN_CACHE_PATH, data, { mode: 0o600 });
}

export async function removeTokenCache(): Promise<void> {
  try {
    await unlink(TOKEN_CACHE_PATH);
  } catch {
    // file may not exist
  }
}
