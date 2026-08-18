import { readFile, writeFile, unlink, rename } from "node:fs/promises";
import { randomBytes } from "node:crypto";
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
  const tmpPath = `${tokensPath}.${randomBytes(4).toString("hex")}.tmp`;
  await writeFile(tmpPath, data, { mode: 0o600 });
  await rename(tmpPath, tokensPath);
}

export async function removeTokenCache(): Promise<void> {
  const { tokensPath } = resolveStoragePaths();
  try {
    await unlink(tokensPath);
  } catch {}
}
