# Design Técnico — office365-mcp-openclaw-isolation

## Contexto de integração

O OpenClaw spawna processos MCP via `npx @leonardocrdso/office365-mcp-server` com env vars por agente:

```json
{
  "command": "npx",
  "args": ["@leonardocrdso/office365-mcp-server"],
  "env": {
    "AZURE_CLIENT_ID": "...",
    "AZURE_TENANT_ID": "...",
    "OFFICE365_MCP_HOME": "/Users/live/.openclaw/agents/gabrielsens/office365"
  }
}
```

Hoje o servidor ignora `OFFICE365_MCP_HOME` e usa `~/` hardcoded. Duas instâncias diferentes sobrescrevem os mesmos arquivos.

---

## Arquivos afetados

| Arquivo | Ação |
|---|---|
| `src/auth/storage.ts` | criar — centraliza resolução de paths |
| `src/auth/config-manager.ts` | editar — importar de storage.ts |
| `src/auth/token-cache-manager.ts` | editar — importar de storage.ts |
| `src/tools/auth-tools.ts` | editar — mostrar path real nas respostas |

---

## Tipos

```typescript
// src/auth/storage.ts
interface StoragePaths {
  configPath: string;
  tokensPath: string;
  baseDir: string | null;  // null = usando homedir legado
}
```

---

## Lógica central — resolveStoragePaths()

```typescript
// src/auth/storage.ts
import { homedir } from "node:os";
import { join } from "node:path";
import { mkdir } from "node:fs/promises";

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
```

---

## config-manager.ts refatorado

```typescript
import { readFile, writeFile } from "node:fs/promises";
import { resolveStoragePaths, ensureStorageDir } from "./storage.js";

export async function loadConfig() {
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
  return configPath;  // retorna o path real para a tool exibir
}
```

`saveConfig` passa a retornar o `configPath` para que a tool `configure` exiba o caminho real.

---

## token-cache-manager.ts refatorado

```typescript
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
```

---

## auth-tools.ts — mensagens atualizadas

**Tool configure:** usar o path retornado por `saveConfig()`:

```typescript
const savedPath = await saveConfig(params.clientId, params.tenantId);
// mensagem:
`As credenciais foram salvas em \`${savedPath}\`.`
```

**Tool auth-status:** incluir path de storage no retorno:

```typescript
const { baseDir, configPath } = resolveStoragePaths();
const storageLabel = baseDir
  ? `\`${baseDir}\` (isolado via OFFICE365_MCP_HOME)`
  : `\`${configPath}\` (storage padrão)`;
// ...
`**Storage:** ${storageLabel}`
```

---

## Regras de backward compat

- `OFFICE365_MCP_HOME` não definida → paths iguais aos de hoje (`~/.office365-mcp-config.json`, `~/.office365-mcp-tokens.json`)
- Nenhuma migração de dados — instâncias legadas continuam funcionando sem tocar os arquivos existentes
- `ensureStorageDir` só cria diretório quando `baseDir !== null` (nunca recria `~/`)

---

## Plano de implementação

**Etapa 1:** criar `src/auth/storage.ts`
**Etapa 2:** refatorar `config-manager.ts` (remove `CONFIG_PATH`, importa de storage)
**Etapa 3:** refatorar `token-cache-manager.ts` (remove `TOKEN_CACHE_PATH`, importa de storage)
**Etapa 4:** atualizar `auth-tools.ts` (configure, auth-status)
**Etapa 5:** build + smoke test
