# @chat-adapter/teams

[![npm version](https://img.shields.io/npm/v/@chat-adapter/teams)](https://www.npmjs.com/package/@chat-adapter/teams)
[![npm downloads](https://img.shields.io/npm/dm/@chat-adapter/teams)](https://www.npmjs.com/package/@chat-adapter/teams)

Microsoft Teams adapter for [Chat SDK](https://chat-sdk.dev).

## Installation

```bash
pnpm add @chat-adapter/teams
```

## Usage

The adapter auto-detects `TEAMS_APP_ID`, `TEAMS_APP_PASSWORD`, and `TEAMS_APP_TENANT_ID` from environment variables:

```typescript
import { Chat } from "chat";
import { createTeamsAdapter } from "@chat-adapter/teams";

const bot = new Chat({
  userName: "mybot",
  adapters: {
    teams: createTeamsAdapter({
      appType: "SingleTenant",
    }),
  },
});

bot.onNewMention(async (thread, message) => {
  await thread.post("Hello from Teams!");
});
```

## Setup with Teams CLI

The fastest way to get started is with the [`teams` CLI](https://heyitsaamir.github.io/teamscli/). It handles app creation, bot registration, credentials, and manifest management from the command line.

### 1. Install and log in

```bash
npm install -g https://github.com/heyitsaamir/teamscli/releases/latest/download/teamscli.tgz
teams login
```

### 2. Create app with bot

```bash
# Creates a Teams-managed app with bot and writes credentials to .env
teams app create --name "My Chat Bot" --endpoint https://your-domain.com/api/webhooks/teams --env .env
```

This creates the app, registers the bot, and writes `TEAMS_APP_ID`, `TEAMS_APP_PASSWORD`, and `TEAMS_APP_TENANT_ID` to your `.env` file.

To create an Azure-hosted bot instead (requires [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/)):

```bash
teams app create --name "My Chat Bot" --endpoint https://your-domain.com/api/webhooks/teams --azure --resource-group my-rg --env .env
```

### 3. Enable receiving all messages (optional)

By default, Teams bots only receive messages when @-mentioned. To receive all messages in channels and group chats, add the RSC permission:

```bash
teams app rsc add <appId> ChannelMessage.Read.Group --type Application
```

### 4. Install in Teams

View the install link for your app:

```bash
teams app view <appId> --web
```

Open the link in your browser to install the app into a team or chat.

See the [Teams CLI docs](https://heyitsaamir.github.io/teamscli/) for the full list of commands.

## Configuration

All options are auto-detected from environment variables when not provided. Internally, the adapter maps these options to the Teams SDK (`@microsoft/teams.apps`).

| Option | Required | Description |
|--------|----------|-------------|
| `appId` | No* | Azure Bot App ID. Auto-detected from `TEAMS_APP_ID` |
| `appPassword` | No** | Azure Bot App Password. Auto-detected from `TEAMS_APP_PASSWORD` |
| `federated` | No** | Federated (workload identity) authentication config |
| `appType` | No | `"MultiTenant"` or `"SingleTenant"` (default: `"MultiTenant"`) |
| `appTenantId` | For SingleTenant | Azure AD Tenant ID. Auto-detected from `TEAMS_APP_TENANT_ID` |
| `userName` | No | Bot display name (default: `"bot"`) |
| `logger` | No | Logger instance (defaults to `ConsoleLogger("info")`) |

\*`appId` is required — either via config or `TEAMS_APP_ID` env var.

\*\*Exactly one authentication method is required: `appPassword` or `federated`. When neither is provided, `TEAMS_APP_PASSWORD` is auto-detected from environment.

### Authentication methods

The adapter supports two authentication methods. When no explicit auth is provided, `TEAMS_APP_PASSWORD` is auto-detected from environment variables.

#### Client secret (default)

The simplest option — provide `appPassword` directly or set `TEAMS_APP_PASSWORD`:

```typescript
createTeamsAdapter({
  appPassword: "your_app_password_here",
});
```

#### Federated (workload identity)

For environments with managed identities (e.g. Azure Kubernetes Service, GitHub Actions). Maps to `managedIdentityClientId` in the Teams SDK:

```typescript
createTeamsAdapter({
  federated: {
    clientId: "your_managed_identity_client_id_here",
  },
});
```

## Environment variables

```bash
TEAMS_APP_ID=...
TEAMS_APP_PASSWORD=...
TEAMS_APP_TENANT_ID=...  # Required for SingleTenant apps
```

## Features

### Messaging

| Feature | Supported |
|---------|-----------|
| Post message | Yes |
| Edit message | Yes |
| Delete message | Yes |
| File uploads | Yes |
| Streaming | Post+Edit fallback |

### Rich content

| Feature | Supported |
|---------|-----------|
| Card format | Adaptive Cards |
| Buttons | Yes |
| Link buttons | Yes |
| Select menus | No |
| Tables | GFM |
| Fields | Yes |
| Images in cards | Yes |
| Modals | No |

### Conversations

| Feature | Supported |
|---------|-----------|
| Slash commands | No |
| Mentions | Yes |
| Add reactions | No |
| Remove reactions | No |
| Receive reactions | Yes |
| Typing indicator | Yes |
| DMs | Yes |
| Ephemeral messages | No (DM fallback) |

### Message history

| Feature | Supported |
|---------|-----------|
| Fetch messages | Yes (requires Graph permissions) |
| Fetch single message | No |
| Fetch thread info | Yes |
| Fetch channel messages | Yes (requires Graph permissions) |
| List threads | Yes (requires Graph permissions) |
| Fetch channel info | Yes (requires Graph permissions) |
| Post channel message | Yes |

## Message history (`fetchMessages`)

Fetching message history requires the Microsoft Graph API with client credentials flow. To enable it:

1. Set `appTenantId` in the adapter config (or `TEAMS_APP_TENANT_ID` env var)
2. Grant one of these Azure AD app permissions:
   - `ChatMessage.Read.Chat`
   - `Chat.Read.All`
   - `Chat.Read.WhereInstalled`

Without these permissions, `fetchMessages` will throw a `NotImplementedError`.

### Receiving all messages

By default, Teams bots only receive messages when directly @-mentioned. To receive all messages in a channel or group chat, add the RSC permission using the CLI:

```bash
teams app rsc add <appId> ChannelMessage.Read.Group --type Application
```

Or add it manually to your Teams app manifest:

```json
{
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "ChannelMessage.Read.Group",
          "type": "Application"
        }
      ]
    }
  }
}
```

## Troubleshooting

Run `teams app doctor <appId>` to automatically diagnose common issues.

### "Unauthorized" error

- Verify `TEAMS_APP_ID` and your chosen auth credential are correct
- For client secret auth, check that `TEAMS_APP_PASSWORD` is valid and not expired
- For federated auth, verify the managed identity client ID is correct and that federated credentials are configured in Azure AD
- For SingleTenant apps, ensure `TEAMS_APP_TENANT_ID` is set
- Check that the messaging endpoint URL is correct in Azure

### Bot not appearing in Teams

- Verify the Teams channel is enabled in Azure Bot
- Check that the app manifest is correctly configured
- Ensure the app is installed in the workspace/team

### Messages not received

- Verify the messaging endpoint URL is correct
- Check that your server is accessible from the internet
- Review Azure Bot logs for errors

## License

MIT
