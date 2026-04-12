# @chat-adapter/teams

[![npm version](https://img.shields.io/npm/v/@chat-adapter/teams)](https://www.npmjs.com/package/@chat-adapter/teams)
[![npm downloads](https://img.shields.io/npm/dm/@chat-adapter/teams)](https://www.npmjs.com/package/@chat-adapter/teams)

Microsoft Teams adapter for [Chat SDK](https://chat-sdk.dev). Configure with Azure Bot Service.

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

## Setup

### Install the Teams CLI

```bash
npm install -g https://github.com/heyitsaamir/teamscli/releases/latest/download/teamscli.tgz
```

### Create and configure the app

```bash
# Log in to Microsoft 365
teams login

# Create app, bot, and credentials in one step
teams app create \
  --name "My Bot" \
  --endpoint "https://your-domain.com/api/webhooks/teams" \
  --env .env

# Receive all messages in channels (not just @mentions)
teams app rsc add <appId> ChannelMessage.Read.Group --type Application

# Receive all messages in group chats
teams app rsc add <appId> ChatMessage.Read.Chat --type Application
```

`teams app create` outputs an installation link — share it to install the app in Teams.

Credentials (`TEAMS_APP_ID`, `TEAMS_APP_PASSWORD`, `TEAMS_APP_TENANT_ID`) are written to `.env` automatically.

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

Fetching message history requires the Microsoft Graph API with client credentials flow. The permissions needed depend on the conversation type:

| Context | Permission | Type |
|---------|-----------|------|
| Channel | `ChannelMessage.Read.Group` | RSC |
| Group chat | `ChatMessage.Read.Chat` | RSC |
| Personal/DM | `Chat.Read.All` | Azure AD (admin consent) |

RSC permissions are added via the Teams CLI (see [setup](#setup)). For DM history, RSC is [not sufficient](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/rsc/resource-specific-consent) — add the `Chat.Read.All` Azure AD app permission:

```bash
az ad app permission add \
  --id <appId> \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions 6b7d71aa-70aa-4810-a8d9-5d9fb2830017=Role

az ad app permission admin-consent --id <appId>
```

Set `appTenantId` in the adapter config (or `TEAMS_APP_TENANT_ID` env var) to enable tenant-scoped Graph tokens.

Without these permissions, `fetchMessages` will throw a `NotImplementedError`.

### Receiving all messages

By default, Teams bots only receive messages when directly @-mentioned. The [setup](#setup) above adds RSC permissions to receive all messages. You can also manage these with the Teams CLI:

```bash
teams app rsc list <appId>
teams app rsc add <appId> ChannelMessage.Read.Group --type Application
teams app rsc add <appId> ChatMessage.Read.Chat --type Application
```

## Troubleshooting

Run diagnostics on your app:

```bash
teams app doctor <appId>
```

## License

MIT
