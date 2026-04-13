import { mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join } from "node:path";

interface StoragePaths {
  configPath: string;
  tokensPath: string;
  baseDir: string | null;
}

export function resolveStoragePaths(): StoragePaths {
  const mcpHome = process.env.OFFICE365_MCP_HOME;
  if (mcpHome) {
    return {
      configPath: join(mcpHome, "config.json"),
      tokensPath: join(mcpHome, "tokens.json"),
      baseDir: mcpHome,
    };
  }
  return {
    configPath: join(homedir(), ".office365-mcp-config.json"),
    tokensPath: join(homedir(), ".office365-mcp-tokens.json"),
    baseDir: null,
  };
}

export async function ensureStorageDir(): Promise<void> {
  const { baseDir } = resolveStoragePaths();
  if (baseDir) {
    await mkdir(baseDir, { recursive: true });
  }
}
