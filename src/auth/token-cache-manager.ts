import { readFile, writeFile, unlink } from "node:fs/promises";
import { resolveStoragePaths, ensureStorageDir } from "./storage.js";

export async function loadTokenCache(): Promise<string | null> {
  const { tokensPath } = resolveStoragePaths();
  try {
    return await readFile(tokensPath, "utf-8");
  } catch {
    return null;
  }
}

export async function saveTokenCache(data: string): Promise<void> {
  await ensureStorageDir();
  const { tokensPath } = resolveStoragePaths();
  await writeFile(tokensPath, data, { mode: 0o600 });
}

export async function removeTokenCache(): Promise<void> {
  const { tokensPath } = resolveStoragePaths();
  try {
    await unlink(tokensPath);
  } catch {}
}
