---
name: outlook-addin
description: Outlook sidebar add-in that brings the full power of your OpenClaw agent into Microsoft Outlook. Chat with your agent about any email, use all your tools and skills, draft replies — directly from the inbox. Works with Outlook Desktop (Classic) and Outlook Web (OWA). Use when a user wants to integrate OpenClaw with Outlook, chat with their agent from email, or set up an AI sidebar in Outlook.
---

# OpenClaw Outlook Add-in

An Office.js sidebar add-in that connects Outlook to your local OpenClaw Gateway via WebSocket. Select any email, and your full agent — with all tools, skills, and automations — is available right in the sidebar.

**Repository:** https://github.com/nachtsheim/openclaw-outlook-addin (MIT)

## What It Does

- Read any email and chat with your OpenClaw agent about it
- Use all your agent's tools directly from Outlook (calendar, trackers, databases, automations — whatever your agent can do)
- One-click draft reply, opens Outlook's native compose for review
- Per-email chat sessions that persist when switching between emails
- Streaming responses with tool call indicators
- Light/dark mode auto-detection

## Requirements

- Node.js 18+
- OpenClaw Gateway running locally
- Microsoft 365 account with sideloading enabled
- Outlook Desktop (Classic, Windows) or Outlook Web (OWA)

## Installation

### 1. Clone and install

```bash
git clone https://github.com/nachtsheim/openclaw-outlook-addin.git
cd openclaw-outlook-addin
npm install
npx office-addin-dev-certs install   # first time only
npm run dev                           # starts https://localhost:3000
```

### 2. Allow the add-in origin in Gateway config

```bash
openclaw config patch '{"gateway":{"controlUi":{"allowedOrigins":["https://localhost:3000"]}}}'
```

### 3. Sideload into Outlook

1. Open https://aka.ms/olksideload (Outlook Web)
2. My add-ins → Add a custom add-in → Add from file → upload `manifest.xml`
3. Sync to Outlook Desktop can take up to 24h

### 4. Connect

Open the sidebar (OpenClaw AI button in ribbon), paste your Gateway token from `~/.openclaw/openclaw.json` → `gateway.auth.token`, click Save & Connect.

## How It Works

- Connects via WebSocket RPC to the local Gateway (proxied through webpack dev server)
- Each email gets its own agent session (keyed by subject + sender + date)
- Email context (subject, sender, body) is sent with the first message per email to save tokens
- Responses stream in real-time with typing indicators during tool calls

## Troubleshooting

- **"Disconnected"** → Check if dev server and Gateway are running
- **No add-in button in Outlook** → Restart Outlook or wait for OWA→Desktop sync
- **Can't sideload** → M365 admin must allow custom add-ins
- **Token prompt every time** → localStorage may be cleared by browser policy

Full troubleshooting guide in the [README](https://github.com/nachtsheim/openclaw-outlook-addin#troubleshooting).
