# OpenClaw Outlook Add-in

AI-powered Outlook sidebar that reads email context, lets you chat with OpenClaw, and draft/send replies — directly from Outlook.

![Status](https://img.shields.io/badge/status-beta-yellow) ![Platform](https://img.shields.io/badge/platform-Outlook%20Desktop%20%7C%20OWA-blue)

## Features

- 📧 **Email Context** — Automatically reads subject, sender, recipients, date, and body of the selected email
- 💬 **Chat Interface** — Ask questions about the email, get summaries, translations, or any AI assistance
- ✏️ **Draft Reply** — One-click reply drafting based on the email context
- 📤 **Send Reply** — Opens Outlook's native reply compose with the drafted text pre-filled
- 🌗 **Light/Dark Mode** — Auto-detects Outlook theme (Office.js + prefers-color-scheme)
- 🔄 **Auto-Reconnect** — WebSocket reconnects automatically with exponential backoff
- 🔒 **Token-based Auth** — Gateway token stored in browser localStorage, never in code

## Prerequisites

- [Node.js](https://nodejs.org/) 18+
- [OpenClaw](https://github.com/openclaw/openclaw) Gateway running locally
- Outlook Desktop (Windows, Classic) or Outlook Web (OWA)
- Microsoft 365 account with sideloading enabled

## Quick Start

### 1. Install dependencies

```bash
npm install
```

### 2. Install dev certificates (first time only)

```bash
npx office-addin-dev-certs install
```

### 3. Start the dev server

```bash
npm run dev
```

Starts webpack-dev-server at `https://localhost:3000` with:
- HTTPS (required by Office.js)
- Hot reload
- WebSocket proxy to OpenClaw Gateway (`/gateway-ws` → `ws://127.0.0.1:18789`)

### 4. Sideload the add-in

**Via OWA (recommended — syncs to Desktop automatically):**

1. Open [https://aka.ms/olksideload](https://aka.ms/olksideload)
2. Click **My add-ins** → **Add a custom add-in** → **Add from file**
3. Upload `manifest.xml` from this project

**Note:** Sideloading in OWA requires your Microsoft 365 admin to allow custom add-ins. Sync to Outlook Desktop can take up to 24 hours.

### 5. Configure Gateway Token

When you first open the add-in sidebar, it will ask for your OpenClaw Gateway token:

1. Find your token in `~/.openclaw/openclaw.json` → `gateway.auth.token`
2. Paste it into the token input field in the sidebar
3. Click **Save & Connect**

The token is stored in browser localStorage and never leaves your machine.

### 6. Gateway Configuration

Add `https://localhost:3000` to your Gateway's allowed origins:

```json
{
  "gateway": {
    "controlUi": {
      "allowedOrigins": ["https://localhost:3000"]
    }
  }
}
```

Or via OpenClaw CLI:
```bash
openclaw config patch '{"gateway":{"controlUi":{"allowedOrigins":["https://localhost:3000"]}}}'
```

## Usage

1. Open any email in Outlook
2. Click the **OpenClaw AI** button in the ribbon (or find it via **...** → More actions)
3. The sidebar shows the email context and a chat interface
4. Type a question or click **Draft Reply**

## Auto-Start (Windows)

To keep the dev server running permanently, create a Windows Scheduled Task:

```powershell
$action = New-ScheduledTaskAction -Execute "node.exe" `
  -Argument "node_modules\webpack-cli\bin\cli.js serve --mode development" `
  -WorkingDirectory "C:\path\to\openclaw-outlook-addin"
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries `
  -DontStopIfGoingOnBatteries -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName "OpenClaw Outlook Add-in" `
  -Action $action -Trigger $trigger -Settings $settings -Force
```

## Build for Production

```bash
npm run build
```

Output goes to `dist/`. For production deployment, host the built files on any HTTPS server and update the URLs in `manifest.xml`.

## Architecture

```
Browser (Outlook WebView)          Local Machine
┌─────────────────────┐           ┌──────────────────────┐
│  Outlook Add-in     │  wss://   │  Webpack Dev Server   │
│  (taskpane.html)    │◄────────►│  :3000 (HTTPS)        │
│                     │           │    │                  │
│  - Office.js        │           │    │ proxy /gateway-ws│
│  - Chat UI          │           │    ▼                  │
│  - Theme detection  │           │  OpenClaw Gateway     │
│                     │           │  :18789 (WS)          │
└─────────────────────┘           └──────────────────────┘
```

## WebSocket Protocol

The add-in uses OpenClaw's native Gateway RPC protocol:

### Authentication

1. Client connects to `wss://localhost:3000/gateway-ws` (proxied to Gateway)
2. Gateway sends `connect.challenge` event
3. Client sends `connect` RPC with auth token
4. Gateway responds with session info

### Sending Messages

```json
{
  "type": "req",
  "id": "1",
  "method": "chat.send",
  "params": {
    "sessionKey": "agent:main:main",
    "message": "Summarize this email",
    "deliver": false
  }
}
```

### Receiving Responses

- **Streaming:** `agent.delta` / `chat.delta` events with incremental text
- **Complete:** `agent.message` / `chat.message` events with full response
- **Tool calls:** `agent.tool_call` / `agent.tool_result` events

## Project Structure

```
├── manifest.xml              # Office Add-in manifest (XML)
├── package.json
├── webpack.config.js         # Dev server + WS proxy config
├── generate-icons.js         # Icon generator (uses sharp + OpenClaw SVG)
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html     # Sidebar UI
│   │   ├── taskpane.js       # Office.js + Gateway RPC + chat logic
│   │   └── taskpane.css      # Light/dark theme styles
│   └── commands/
│       ├── commands.html
│       └── commands.js       # Ribbon command handlers
└── assets/
    ├── openclaw-logo.svg     # Source logo (from Gateway favicon)
    └── icon-*.png            # Generated icons (16-128px)
```

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "Add-in installation failed" | Validate manifest: `npx office-addin-manifest validate manifest.xml` |
| "Disconnected" in sidebar | Check if dev server is running (`npm run dev`) |
| "Connecting..." stays | Verify Gateway is running and token is correct |
| No button in Outlook ribbon | Restart Outlook, or wait for OWA→Desktop sync |
| Icons not showing | Remove add-in in OWA, re-sideload `manifest.xml` |
| Mixed content errors | The WSS proxy should handle this — check webpack proxy config |
| Need to debug the sidebar | Open [https://localhost:3000/taskpane.html](https://localhost:3000/taskpane.html) directly in a browser tab |

## Security Notes

- Gateway token is stored in **browser localStorage only** — never committed to Git
- The add-in has **ReadItem** permission (read-only access to current email)
- `displayReplyForm()` opens Outlook's native compose — user always reviews before sending
- WebSocket connection is local only (`localhost:3000` → `localhost:18789`)

## Privacy & Data Protection (GDPR / DSGVO)

⚠️ **Important:** This add-in sends email content (subject, sender, recipients, body text) to an AI model via your local OpenClaw Gateway. Before using it in a professional context, ensure the following:

- **Data processing agreement (DPA/AVV):** If you process personal data from emails, ensure your AI provider has an appropriate data processing agreement in place.
- **Zero data retention (ZDR):** Use an AI provider plan that does **not** retain or train on your data (e.g., Anthropic API, OpenAI API with ZDR, Google Gemini API — **not** free-tier consumer products).
- **Local processing:** All data flows through your **local** OpenClaw Gateway (`localhost`). No data is sent to third-party servers other than the configured AI model provider.
- **No cloud storage:** The add-in does not store emails, conversations, or tokens on any external server. The Gateway token is stored in browser localStorage on your machine only.
- **Employee consent:** If processing employee or customer emails, ensure appropriate legal basis under GDPR Art. 6 (e.g., legitimate interest, consent, or contractual necessity).
- **Data minimization:** The add-in sends only the currently selected email's content — not your entire mailbox.

**Recommendation:** For business use, pair this add-in with an enterprise AI plan that provides contractual guarantees for data privacy, such as:
- Anthropic Claude (API / Max for Business) — zero retention by default
- OpenAI API (with Zero Data Retention) — opt-in via API settings
- Self-hosted models (Ollama, vLLM) — data never leaves your infrastructure

**This add-in is a tool — compliance with data protection regulations is the responsibility of the deploying organization.**

## License

MIT
