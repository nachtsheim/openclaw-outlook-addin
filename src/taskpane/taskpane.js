import "./taskpane.css";

// ===== Configuration =====
const GATEWAY_PORT = 18789;
const RECONNECT_BASE_DELAY = 1000;
const RECONNECT_MAX_DELAY = 15000;

// ===== State =====
let ws = null;
let reconnectAttempts = 0;
let reconnectTimer = null;
let connected = false;
let sessionKey = "agent:main:outlook-addin";
let lastDraftContent = null;
let emailContext = null;
let rpcId = 0;
let pendingRpc = new Map();
let currentStream = "";
let currentRunId = null;
let gatewayToken = null;
let waitingForResponse = false;
let historyFetchPending = false;

// ===== DOM References =====
const $ = (id) => document.getElementById(id);

// ===== Office.js Initialization =====
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeAddin();
  }
});

function initializeAddin() {
  detectAndApplyTheme();
  bindUIEvents();
  readGatewayToken();
  readEmailContext();
  // Only connect if we have a token
  if (gatewayToken) {
    connectWebSocket();
  }

  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => readEmailContext(),
      () => {}
    );
  }
}

// ===== Theme Detection =====
function detectAndApplyTheme() {
  let isDark = false;
  try {
    const theme = Office.context.officeTheme;
    if (theme && theme.bodyBackgroundColor) {
      const bg = theme.bodyBackgroundColor.replace("#", "");
      const r = parseInt(bg.substring(0, 2), 16);
      const g = parseInt(bg.substring(2, 4), 16);
      const b = parseInt(bg.substring(4, 6), 16);
      isDark = (0.299 * r + 0.587 * g + 0.114 * b) / 255 < 0.5;
    }
  } catch (e) {}

  if (!isDark && window.matchMedia) {
    isDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", (e) => {
      document.documentElement.setAttribute("data-theme", e.matches ? "dark" : "light");
    });
  }
  document.documentElement.setAttribute("data-theme", isDark ? "dark" : "light");
}

// ===== Gateway Token =====
function readGatewayToken() {
  try {
    gatewayToken = localStorage.getItem("openclaw-gateway-token") || null;
  } catch (e) {}

  // If no token stored, show settings prompt
  if (!gatewayToken) {
    showTokenSetup();
  }
}

function showTokenSetup() {
  const container = $("chat-messages");
  const div = document.createElement("div");
  div.className = "message system-message";
  div.innerHTML = `<div class="message-content">
    <strong>Gateway Token Required</strong><br>
    Enter your OpenClaw Gateway token to connect:<br><br>
    <input type="password" id="token-input" placeholder="Paste gateway token..." 
      style="width:100%;padding:6px 8px;border:1px solid var(--border);border-radius:4px;background:var(--bg-input);color:var(--text-primary);font-size:12px;margin-bottom:6px;font-family:var(--font-mono)"/>
    <button id="token-save-btn" 
      style="padding:5px 12px;background:var(--accent);color:#fff;border:none;border-radius:4px;cursor:pointer;font-size:12px">
      Save & Connect
    </button>
    <br><small style="color:var(--text-muted)">Find it in ~/.openclaw/openclaw.json → gateway.auth.token</small>
  </div>`;
  container.appendChild(div);

  setTimeout(() => {
    const btn = document.getElementById("token-save-btn");
    const input = document.getElementById("token-input");
    if (btn && input) {
      btn.addEventListener("click", () => {
        const token = input.value.trim();
        if (token) {
          try { localStorage.setItem("openclaw-gateway-token", token); } catch(e) {}
          gatewayToken = token;
          div.remove();
          addMessage("system", "Token saved. Connecting...");
          connectWebSocket();
        }
      });
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") btn.click();
      });
    }
  }, 100);
}

// ===== Email Context =====
function readEmailContext() {
  const item = Office.context.mailbox.item;
  if (!item) { showEmailPlaceholder(); return; }

  try {
    const subject = item.subject || "(No subject)";
    const from = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "Unknown";
    const dateTime = item.dateTimeCreated ? new Date(item.dateTimeCreated).toLocaleString() : "";
    const to = item.to ? item.to.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ") : "";
    const cc = item.cc ? item.cc.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ") : "";

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      const body = result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
      emailContext = { subject, from, to, cc, date: dateTime, body };
      showEmailInfo(subject, from, dateTime);
    });
  } catch (err) {
    showEmailPlaceholder();
    addMessage("error", "Failed to read email.");
  }
}

function showEmailPlaceholder() {
  $("email-placeholder").style.display = "flex";
  $("email-info").style.display = "none";
  emailContext = null;
}

function showEmailInfo(subject, from, date) {
  $("email-placeholder").style.display = "none";
  $("email-info").style.display = "block";
  $("email-subject").textContent = subject;
  $("email-from").textContent = from;
  $("email-date").textContent = date;
}

// ===== WebSocket (OpenClaw Gateway Protocol) =====
function getWsUrl() {
  // Use the webpack dev-server proxy to avoid mixed-content (https → ws) blocking
  // The proxy at wss://localhost:3000/gateway-ws forwards to ws://127.0.0.1:18789
  const loc = window.location;
  if (loc.protocol === "https:") {
    return `wss://${loc.host}/gateway-ws`;
  }
  return `ws://127.0.0.1:${GATEWAY_PORT}`;
}

function connectWebSocket() {
  if (ws && (ws.readyState === WebSocket.OPEN || ws.readyState === WebSocket.CONNECTING)) return;

  setConnectionStatus("connecting");
  
  try {
    ws = new WebSocket(getWsUrl());
  } catch (e) {
    setConnectionStatus("disconnected");
    scheduleReconnect();
    return;
  }

  ws.onopen = () => {
    // WS transport is open but Gateway handshake not done yet — stay on "connecting"
    setConnectionStatus("connecting");
    // Wait for connect.challenge event from server, then send connect
    // If no challenge comes within 2s, send connect anyway
    setTimeout(() => {
      if (!connected) sendConnect();
    }, 2000);
  };

  ws.onmessage = (event) => {
    handleMessage(String(event.data || ""));
  };

  ws.onclose = (event) => {
    connected = false;
    pendingRpc.clear();
    setConnectionStatus("disconnected");
    scheduleReconnect();
  };

  ws.onerror = () => {};
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  const delay = Math.min(RECONNECT_BASE_DELAY * Math.pow(1.7, reconnectAttempts), RECONNECT_MAX_DELAY);
  reconnectAttempts++;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectWebSocket();
  }, delay);
}

function sendConnect() {
  const params = {
    minProtocol: 3,
    maxProtocol: 3,
    client: {
      id: "openclaw-control-ui",
      version: "1.0.0",
      platform: navigator.platform || "web",
      mode: "webchat",
      instanceId: "outlook-" + Date.now()
    },
    role: "operator",
    scopes: ["operator.admin"],
    caps: ["tool-events"],
    auth: {}
  };

  if (gatewayToken) {
    params.auth = { token: gatewayToken };
  }

  rpcRequest("connect", params)
    .then((result) => {
      connected = true;
      reconnectAttempts = 0;
      setConnectionStatus("connected");
      
      // Get session key from connect response
      if (result && result.sessionKey) {
        sessionKey = result.sessionKey;
      }
      console.log("[openclaw] connected, session:", sessionKey);
    })
    .catch((err) => {
      console.error("Connect failed:", err);
      setConnectionStatus("disconnected");
    });
}

function rpcRequest(method, params) {
  return new Promise((resolve, reject) => {
    if (!ws || ws.readyState !== WebSocket.OPEN) {
      reject(new Error("WebSocket not connected"));
      return;
    }
    const id = String(++rpcId);
    pendingRpc.set(id, { resolve, reject });
    ws.send(JSON.stringify({ type: "req", id, method, params }));
  });
}

function handleMessage(raw) {
  let data;
  try { data = JSON.parse(raw); } catch { return; }

  // RPC Response (type: "res")
  if (data.type === "res" && pendingRpc.has(String(data.id))) {
    const { resolve, reject } = pendingRpc.get(String(data.id));
    pendingRpc.delete(String(data.id));
    if (data.ok === false) {
      reject(new Error(data.error?.message || data.error?.code || "RPC error"));
    } else {
      // Gateway returns result in either data.result or data.payload
      resolve(data.result || data.payload || data);
    }
    return;
  }

  // Event (server push)
  if (data.type === "event") {
    handleEvent(data);
    return;
  }
}

function handleEvent(evt) {
  const event = evt.event || "";
  const payload = evt.payload || evt.data || {};
  // Debug: log all events
  console.log("[openclaw] event:", event, JSON.stringify(payload).substring(0, 200));

  switch (event) {
    case "connect.challenge":
      sendConnect();
      break;

    case "agent.run": {
      const phase = payload.phase || payload.data?.phase || "";
      if (phase === "start") {
        currentRunId = payload.runId || null;
        currentStream = "";
        showTyping();
      } else if (phase === "end" || phase === "error") {
        currentRunId = null;
        hideTyping();
        if (currentStream.trim()) {
          addMessage("ai", currentStream.trim());
          currentStream = "";
        }
      }
      break;
    }

    case "chat": {
      const state = payload.state || "";
      if (state === "start" || state === "started") {
        currentRunId = payload.runId || null;
        currentStream = "";
        showTyping();
      } else if (state === "final" || state === "end" || state === "error") {
        currentRunId = null;
        hideTyping();
        if (currentStream.trim()) {
          addMessage("ai", currentStream.trim());
          currentStream = "";
        } else if (waitingForResponse) {
          waitingForResponse = false;
          fetchLastAssistantMessage();
        }
      }
      break;
    }

    case "agent.delta":
    case "chat.delta": {
      const text = payload.delta || payload.text || payload.content || "";
      if (text) {
        currentStream += text;
        updateStreamingMessage(currentStream);
      }
      break;
    }

    case "agent.message":
    case "chat.message": {
      hideTyping();
      const content = payload.content || payload.text || payload.message || "";
      if (content) {
        currentStream = "";
        addMessage("ai", typeof content === "string" ? content : JSON.stringify(content));
      }
      break;
    }

    case "agent.tool_call":
    case "agent.tool_result":
      // Show tool activity indicator
      if (payload.name || payload.toolName) {
        updateTypingText(`Using ${payload.name || payload.toolName}...`);
      }
      break;

    case "session.update":
      // Session metadata update, ignore
      break;

    default:
      // Check if it has content we should display
      if (payload.content || payload.text || payload.message) {
        const text = payload.content || payload.text || payload.message;
        if (typeof text === "string" && text.trim()) {
          hideTyping();
          addMessage("ai", text.trim());
        }
      }
      break;
  }
}

function fetchLastAssistantMessage() {
  if (historyFetchPending) return;
  historyFetchPending = true;
  rpcRequest("chat.history", { sessionKey: sessionKey, limit: 10 })
    .then((result) => {
      if (!result) { addMessage("error", "Empty history response"); return; }
      
      // Handle various response formats
      let messages = [];
      if (Array.isArray(result.messages)) messages = result.messages;
      else if (Array.isArray(result)) messages = result;
      else if (result.history && Array.isArray(result.history)) messages = result.history;
      else {
        // Try to find messages anywhere in the result
        for (const key of Object.keys(result)) {
          if (Array.isArray(result[key])) { messages = result[key]; break; }
        }
      }

      if (messages.length === 0) {
        historyFetchPending = false;
        setTimeout(() => fetchLastAssistantMessage(), 2000);
        return;
      }

      // Find the last assistant message
      for (let i = messages.length - 1; i >= 0; i--) {
        const msg = messages[i];
        if (msg.role === "assistant") {
          let text = "";
          if (typeof msg.content === "string") {
            text = msg.content;
          } else if (Array.isArray(msg.content)) {
            text = msg.content
              .filter(c => c.type === "text")
              .map(c => c.text || "")
              .join("\n");
          }
          if (text.trim()) {
            historyFetchPending = false;
            addMessage("ai", text.trim());
          }
          return;
        }
      }
      // No assistant message yet, retry
      historyFetchPending = false;
      setTimeout(() => fetchLastAssistantMessage(), 2000);
    })
    .catch((err) => {
      historyFetchPending = false;
      addMessage("error", "Failed to fetch response: " + err.message);
    });
}

function setConnectionStatus(status) {
  const bar = $("connection-bar");
  const text = $("connection-text");
  bar.className = "connection-bar " + status;
  const labels = {
    connected: "Connected to OpenClaw",
    connecting: "Connecting...",
    disconnected: "Disconnected",
  };
  text.textContent = labels[status] || status;
}

// ===== Send Message via Gateway Protocol =====
function sendChatMessage(message) {
  if (!connected) {
    addMessage("error", "Not connected to OpenClaw. Reconnecting...");
    connectWebSocket();
    return;
  }

  // Build the message with email context
  let fullMessage = message;
  if (emailContext && !message.startsWith("[Email context already provided]")) {
    const body = (emailContext.body || "").substring(0, 3000);
    fullMessage = `[Current email context]\nSubject: ${emailContext.subject || ""}\nFrom: ${emailContext.from || ""}\nTo: ${emailContext.to || ""}\nDate: ${emailContext.date || ""}\n\nBody:\n${body}\n\n---\n\nUser question: ${message}`;
  }

  rpcRequest("chat.send", {
    sessionKey: sessionKey,
    message: fullMessage,
    deliver: false,
    idempotencyKey: crypto.randomUUID()
  }).then((result) => {
    console.log("[openclaw] chat.send result:", JSON.stringify(result || {}).substring(0, 300));
  }).catch((err) => {
    hideTyping();
    addMessage("error", "Failed to send: " + err.message);
  });
}

// ===== Chat UI =====
function addMessage(role, text) {
  // Remove streaming message if exists
  const existingStream = document.querySelector(".streaming-message");
  if (existingStream) existingStream.remove();

  const container = $("chat-messages");
  const div = document.createElement("div");
  const classMap = {
    user: "message user-message",
    ai: "message ai-message",
    system: "message system-message",
    error: "message error-message",
  };
  div.className = classMap[role] || "message system-message";

  const content = document.createElement("div");
  content.className = "message-content";
  content.textContent = text;
  div.appendChild(content);

  container.appendChild(div);
  scrollToBottom();
}

function updateStreamingMessage(text) {
  hideTyping();
  let el = document.querySelector(".streaming-message");
  if (!el) {
    const container = $("chat-messages");
    el = document.createElement("div");
    el.className = "message ai-message streaming-message";
    const content = document.createElement("div");
    content.className = "message-content";
    el.appendChild(content);
    container.appendChild(el);
  }
  el.querySelector(".message-content").textContent = text;
  scrollToBottom();
}

function showTyping() {
  $("typing-indicator").style.display = "flex";
  scrollToBottom();
}

function hideTyping() {
  $("typing-indicator").style.display = "none";
}

function updateTypingText(text) {
  const el = document.querySelector(".typing-text");
  if (el) el.textContent = text;
  showTyping();
}

function scrollToBottom() {
  const container = $("chat-messages");
  requestAnimationFrame(() => { container.scrollTop = container.scrollHeight; });
}

// ===== User Input =====
function handleSend() {
  const input = $("message-input");
  const text = input.value.trim();
  if (!text) return;

  addMessage("user", text);
  input.value = "";
  autoResizeInput(input);
  showTyping();
  waitingForResponse = true;
  sendChatMessage(text);
}

function autoResizeInput(el) {
  el.style.height = "auto";
  el.style.height = Math.min(el.scrollHeight, 100) + "px";
}

// ===== Draft Reply =====
function handleDraftReply() {
  if (!emailContext) {
    addMessage("error", "No email selected.");
    return;
  }
  addMessage("user", "Draft a reply to this email");
  showTyping();
  sendChatMessage("Please draft a professional reply to this email. Respond in the same language as the original email.");
}

function handleSendReply() {
  // Find the last AI message content
  const aiMessages = document.querySelectorAll(".ai-message .message-content");
  const lastAi = aiMessages[aiMessages.length - 1];
  if (!lastAi) {
    addMessage("error", "No draft available. Click 'Draft Reply' first.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) { addMessage("error", "No email selected."); return; }

  try {
    item.displayReplyForm(lastAi.textContent);
    addMessage("system", "Reply draft opened in Outlook. Review and send.");
  } catch (err) {
    addMessage("error", "Failed to open reply: " + err.message);
  }
}

// ===== Event Binding =====
function bindUIEvents() {
  $("send-btn").addEventListener("click", handleSend);
  $("draft-btn").addEventListener("click", handleDraftReply);
  $("send-reply-btn").addEventListener("click", handleSendReply);

  const input = $("message-input");
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); }
  });
  input.addEventListener("input", () => autoResizeInput(input));
}
