# Office 365 MCP Server

MCP (Model Context Protocol) server for Microsoft 365 integration. Allows AI assistants like Claude to interact with Outlook, Calendar, OneDrive, SharePoint and Teams via Microsoft Graph API.

## Features

**30 tools** across 6 modules:

| Module | Tools | Description |
|--------|-------|-------------|
| **Auth** | `login`, `auth-status`, `logout` | Device Code Flow authentication |
| **Mail** | `list-emails`, `search-emails`, `read-email`, `send-email`, `reply-email`, `list-mail-folders` | Outlook email operations |
| **Calendar** | `list-events`, `create-event`, `update-event`, `delete-event`, `find-free-slots` | Calendar & meeting management |
| **OneDrive** | `list-drive-files`, `read-file-content`, `upload-file`, `search-files`, `share-file` | File storage operations |
| **SharePoint** | `list-sites`, `get-site`, `list-document-libraries`, `list-library-items`, `search-sharepoint` | SharePoint site & document management |
| **Teams** | `list-teams`, `list-channels`, `list-channel-messages`, `send-channel-message`, `list-chats`, `send-chat-message` | Teams messaging |

## Prerequisites

- [Bun](https://bun.sh/) runtime
- Azure AD (Entra ID) app registration

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
5. Copy the **Application (client) ID**

## Installation

```bash
git clone https://github.com/leonardocrdso/office365-mcp-server.git
cd office365-mcp-server
bun install
```

## Configuration

Set environment variables:

```bash
export AZURE_CLIENT_ID=your-client-id
export AZURE_TENANT_ID=common  # or your specific tenant ID
```

## Usage

### With Claude Code

```bash
claude mcp add office365 -e AZURE_CLIENT_ID=your-client-id -- bun run /path/to/office365-mcp-server/src/index.ts
```

### With Claude Desktop

Add to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office365": {
      "command": "bun",
      "args": ["run", "/path/to/office365-mcp-server/src/index.ts"],
      "env": {
        "AZURE_CLIENT_ID": "your-client-id",
        "AZURE_TENANT_ID": "common"
      }
    }
  }
}
```

### Authentication

After adding the server, use the `login` tool. It will return a code and URL — open the URL in your browser, enter the code and sign in with your Microsoft account. Tokens are cached locally at `~/.office365-mcp-tokens.json`.

## Tech Stack

- **Runtime:** [Bun](https://bun.sh/)
- **Protocol:** [MCP SDK](https://github.com/modelcontextprotocol/typescript-sdk) v1.26
- **Auth:** [MSAL Node](https://github.com/AzureAD/microsoft-authentication-library-for-js) (Device Code Flow)
- **API:** [Microsoft Graph](https://learn.microsoft.com/en-us/graph/overview) v1.0

## License

[MIT](LICENSE)
