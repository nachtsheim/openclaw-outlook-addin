---
name: outlook-addin
description: Outlook sidebar add-in that connects to the local OpenClaw Gateway via WebSocket, giving full agent access from within Outlook. Use when setting up, configuring, or troubleshooting the OpenClaw Outlook Add-in. Triggers on Outlook add-in setup, Outlook sidebar, email context chat, Office.js sideloading, or Outlook Gateway integration.
---

# Outlook Add-in

Office.js sidebar add-in for Outlook Desktop (Classic) and OWA that connects to the local OpenClaw Gateway. Reads email context (subject, sender, body) and provides a chat interface to the full agent — all tools, skills, and automations available directly from the inbox.

## Setup

### Prerequisites
- Node.js 18+
- OpenClaw Gateway running locally (port 18789)
- Microsoft 365 account with sideloading enabled

### Install & Run

```bash
cd <project-dir>
npm install
npx office-addin-dev-certs install   # first time only
npm run dev                           # starts https://localhost:3000
```

### Sideload

1. Open https://aka.ms/olksideload (OWA)
2. My add-ins → Add a custom add-in → Add from file → upload `manifest.xml`
3. OWA → Desktop sync can take up to 24h

### Gateway Config

Add localhost:3000 to allowed origins:

```
gateway.controlUi.allowedOrigins: ["https://localhost:3000"]
```

### Token

On first open, paste the Gateway token from `~/.openclaw/openclaw.json` → `gateway.auth.token` into the settings panel (⚙️).

## Architecture

- **Protocol:** WebSocket RPC to Gateway (`/gateway-ws` proxied via webpack-dev-server)
- **Client-ID:** `openclaw-control-ui` (operator.admin scope)
- **Sessions:** Per-email sessions via hash of subject+from+date (`agent:main:outlook-{hash}`)
- **Context:** Email body sent only with first message per email (token-efficient)
- **Streaming:** `agent.delta`/`chat.delta` events rendered incrementally, multi-segment turns flush between tool calls

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Redirect to login on sideload | Admin must allow custom add-ins in M365 |
| "Disconnected" | Check dev server (`npm run dev`) and Gateway status |
| No streaming, only final message | Verify WSS proxy in webpack.config.js (`/gateway-ws`) |
| Duplicate responses | Kill stale node processes on port 3000 |
| Icons missing after update | Remove + re-sideload manifest in OWA |

## Key Files

- `manifest.xml` — Office Add-in manifest (sideloading)
- `src/taskpane/taskpane.js` — Gateway RPC, chat logic, Office.js integration
- `src/taskpane/taskpane.css` — Light/dark theme styles
- `webpack.config.js` — Dev server + WSS proxy config
