# SPEC-001: office365-mcp-openclaw-isolation

**Status:** ready
**Created:** 2026-04-13T14:58:55.037Z
**Updated:** 2026-04-13T15:00:10.090Z

## Overview
Suporte a storage isolado por agente no office365-mcp-server via env var OFFICE365_MCP_HOME. Permite que o OpenClaw rode múltiplas instâncias deste servidor com auth completamente separada por agente.

## Requirements

### REQ-001.1: Centralizar resolução de paths em resolveStoragePaths()

### REQ-001.2: Suporte a OFFICE365_MCP_HOME para config e token cache isolados

### REQ-001.3: Criar diretório OFFICE365_MCP_HOME automaticamente se não existir

### REQ-001.4: Compatibilidade retroativa com paths legados quando env não definida

### REQ-001.5: Tool configure exibir path real onde config foi salva

### REQ-001.6: Tool auth-status exibir path de storage em uso

### REQ-001.7: Garantir que logout use o mesmo storage resolvido

## Acceptance Criteria

- [ ] REQ-001.1: Centralizar resolução de paths em resolveStoragePaths()
- [ ] REQ-001.2: Suporte a OFFICE365_MCP_HOME para config e token cache isolados
- [ ] REQ-001.3: Criar diretório OFFICE365_MCP_HOME automaticamente se não existir
- [ ] REQ-001.4: Compatibilidade retroativa com paths legados quando env não definida
- [ ] REQ-001.5: Tool configure exibir path real onde config foi salva
- [ ] REQ-001.6: Tool auth-status exibir path de storage em uso
- [ ] REQ-001.7: Garantir que logout use o mesmo storage resolvido

## Out of Scope
