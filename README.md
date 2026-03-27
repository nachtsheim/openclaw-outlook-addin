# OpenClaw Outlook Add-in

AI-powered Outlook sidebar that reads email context, lets you chat with OpenClaw, and draft/send replies.

## Prerequisites

- Node.js 18+
- Outlook Desktop (Windows/Mac) or Outlook Web (OWA)
- OpenClaw Gateway running at `ws://localhost:18789`

## Setup

```bash
npm install
```

## Development

### 1. Install dev certificates (first time only)

```bash
npx office-addin-dev-certs install
```

### 2. Start the dev server

```bash
npm run dev
```

This starts a webpack-dev-server at `https://localhost:3000` with hot reload.

### 3. Sideload the add-in

#### Outlook Web (OWA)

1. Go to [Outlook on the web](https://outlook.office.com)
2. Open any email
3. Click **...** (More actions) > **Get Add-ins**
4. Click **My add-ins** > **Add a custom add-in** > **Add from file**
5. Upload `manifest.xml` from this project

#### Outlook Desktop (Windows)

1. Open Outlook
2. Go to **File** > **Manage Add-ins** (opens browser)
3. Click **My add-ins** > **Add a custom add-in** > **Add from file**
4. Upload `manifest.xml`

#### Outlook Desktop (via admin sideload)

1. Place `manifest.xml` in the network share or local folder:
   `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
2. Restart Outlook

### 4. Use the add-in

1. Open any email in Outlook
2. Click the **OpenClaw Assistant** button in the ribbon (or find it under add-ins)
3. The sidebar shows the email subject/sender and a chat interface
4. Type a question or click **Draft Reply**

## Build for Production

```bash
npm run build
```

Output goes to `dist/`.

## WebSocket Protocol

The add-in connects to `ws://localhost:18789` and exchanges JSON messages:

### Client → Server

```json
{
  "type": "message",
  "content": "user's question here",
  "context": {
    "subject": "Email subject",
    "from": "Sender Name <sender@example.com>",
    "to": "Recipient <recipient@example.com>",
    "cc": "",
    "date": "3/27/2026, 10:00:00 AM",
    "body": "Plain text email body..."
  }
}
```

Draft requests use `"type": "draft_request"`.

### Server → Client

```json
{ "type": "message", "content": "AI response text" }
{ "type": "draft", "content": "Drafted reply text" }
{ "type": "error", "content": "Error description" }
```

## Project Structure

```
├── manifest.xml          # Office Add-in manifest
├── package.json
├── webpack.config.js
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html # Sidebar UI
│   │   ├── taskpane.js   # Office.js + WebSocket + chat logic
│   │   └── taskpane.css  # Dark theme styles
│   └── commands/
│       ├── commands.html
│       └── commands.js   # Ribbon command handlers (placeholder)
└── assets/               # Add-in icons (add your own)
```

## Notes

- The add-in requires `ReadItem` permission (read-only access to the current email)
- `displayReplyForm()` opens Outlook's native reply compose window with the drafted text pre-filled — the user always reviews before sending
- WebSocket reconnects automatically with exponential backoff (1s → 30s max)
