# @leonardocrdso/office365-mcp-server

An MCP (Model Context Protocol) server for Microsoft 365 integration. Interact with Outlook, Calendar, OneDrive, SharePoint and Teams via Microsoft Graph API.

## Installation

### Using npx (recommended)

```bash
AZURE_CLIENT_ID=your-client-id npx @leonardocrdso/office365-mcp-server
```

### Adding to Claude Code

```bash
claude mcp add office365-mcp-server -e AZURE_CLIENT_ID=your-client-id -- npx @leonardocrdso/office365-mcp-server
```

### Manual configuration

Add to your MCP client settings:

```json
{
  "mcpServers": {
    "office365-mcp-server": {
      "command": "npx",
      "args": ["@leonardocrdso/office365-mcp-server"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "common"
      }
    }
  }
}
```

## Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations**
2. Click **New registration**:
   - Name: `Office 365 MCP Server`
   - Supported account types: **Accounts in any organizational directory** (multi-tenant)
   - Redirect URI: leave empty
3. In **Authentication**: set **Allow public client flows** to **Yes**
4. In **API permissions** > **Add permissions** > **Microsoft Graph** > **Delegated permissions**, add:
   - `User.Read`
   - `Mail.Read`, `Mail.Send`, `Mail.ReadWrite`
   - `Calendars.Read`, `Calendars.ReadWrite`
   - `Files.Read.All`, `Files.ReadWrite.All`
   - `Sites.Read.All`, `Sites.ReadWrite.All`
   - `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `ChannelMessage.Send`
   - `Chat.Read`, `Chat.ReadWrite`
5. Copy the **Application (client) ID** and use as `AZURE_CLIENT_ID`

## Tools

### Auth

| Tool | Description |
|------|-------------|
| `configure` | Set Azure credentials (Client ID and Tenant ID) — saved to persistent storage |
| `login` | Start Device Code Flow authentication — returns a code and URL |
| `auth-status` | Check authentication status, logged-in user info and active storage path |
| `logout` | Sign out and remove cached tokens |

### Mail

| Tool | Description |
|------|-------------|
| `list-emails` | List emails from Outlook (inbox or specific folder) |
| `search-emails` | Search emails by KQL query |
| `read-email` | Read full email content by ID |
| `send-email` | Send a new email |
| `reply-email` | Reply to an existing email |
| `list-mail-folders` | List email folders |

### Calendar

| Tool | Description |
|------|-------------|
| `list-events` | List calendar events in a date range |
| `create-event` | Create a new event or meeting |
| `update-event` | Update an existing event |
| `delete-event` | Delete an event |
| `find-free-slots` | Check availability for meeting attendees |

### OneDrive

| Tool | Description |
|------|-------------|
| `list-drive-files` | List files and folders |
| `read-file-content` | Read text file content |
| `upload-file` | Upload a file (up to 4MB text) |
| `search-files` | Search files by text |
| `share-file` | Create a sharing link |

### SharePoint

| Tool | Description |
|------|-------------|
| `list-sites` | List SharePoint sites |
| `get-site` | Get site details |
| `list-document-libraries` | List document libraries |
| `list-library-items` | List library items |
| `search-sharepoint` | Search SharePoint content |

### Teams

| Tool | Description |
|------|-------------|
| `list-teams` | List joined teams |
| `list-channels` | List team channels |
| `list-channel-messages` | List channel messages |
| `send-channel-message` | Send a channel message |
| `list-chats` | List direct chats |
| `send-chat-message` | Send a chat message |

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `AZURE_CLIENT_ID` | Yes | — | App registration client ID |
| `AZURE_TENANT_ID` | No | `common` | Tenant ID or `common` for multi-tenant |
| `OFFICE365_MCP_HOME` | No | `~/` | Directory for isolated config and token storage |

### Isolated storage (multi-agent)

When `OFFICE365_MCP_HOME` is set, config and token cache are stored inside that directory (`config.json` and `tokens.json`). The directory is created automatically if it doesn't exist.

This allows running multiple isolated instances of the server — for example, one per agent in an orchestrator like OpenClaw — without auth leaking between them:

```json
{
  "mcpServers": {
    "office365-agent-a": {
      "command": "npx",
      "args": ["@leonardocrdso/office365-mcp-server"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id",
        "OFFICE365_MCP_HOME": "/path/to/agents/agent-a/office365"
      }
    },
    "office365-agent-b": {
      "command": "npx",
      "args": ["@leonardocrdso/office365-mcp-server"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "your-tenant-id",
        "OFFICE365_MCP_HOME": "/path/to/agents/agent-b/office365"
      }
    }
  }
}
```

Without `OFFICE365_MCP_HOME`, storage falls back to `~/.office365-mcp-config.json` and `~/.office365-mcp-tokens.json` (default behavior, unchanged).

## Development

```bash
# Install dependencies
bun install

# Run in development mode
bun run dev

# Build
bun run build
```

## License

MIT
