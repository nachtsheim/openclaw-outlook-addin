import "./taskpane.css";

// ===== Configuration =====
const WS_URL = "ws://localhost:18789";
const RECONNECT_BASE_DELAY = 1000;
const RECONNECT_MAX_DELAY = 30000;
const PING_INTERVAL = 30000;

// ===== State =====
let ws = null;
let reconnectAttempts = 0;
let reconnectTimer = null;
let pingTimer = null;
let lastDraftContent = null;
let emailContext = null;

// ===== DOM References =====
const $ = (id) => document.getElementById(id);

// ===== Office.js Initialization =====
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeAddin();
  }
});

function initializeAddin() {
  bindUIEvents();
  connectWebSocket();
  readEmailContext();

  // Listen for item selection changes
  if (Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      onItemChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.warn("Could not register ItemChanged handler:", result.error.message);
        }
      }
    );
  }
}

// ===== Email Context =====
function readEmailContext() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showEmailPlaceholder();
    return;
  }

  try {
    const subject = item.subject || "(No subject)";
    const from = item.from
      ? `${item.from.displayName} <${item.from.emailAddress}>`
      : "Unknown sender";
    const dateTime = item.dateTimeCreated
      ? new Date(item.dateTimeCreated).toLocaleString()
      : "";
    const to = item.to
      ? item.to.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ")
      : "";
    const cc = item.cc
      ? item.cc.map((r) => `${r.displayName} <${r.emailAddress}>`).join(", ")
      : "";

    // Read the body text
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      const body = result.status === Office.AsyncResultStatus.Succeeded
        ? result.value
        : "";

      emailContext = { subject, from, to, cc, date: dateTime, body };
      showEmailInfo(subject, from, dateTime);
    });
  } catch (err) {
    console.error("Failed to read email context:", err);
    showEmailPlaceholder();
    addMessage("error", "Failed to read email. Please reopen the add-in.");
  }
}

function onItemChanged() {
  readEmailContext();
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

// ===== WebSocket =====
function connectWebSocket() {
  if (ws && (ws.readyState === WebSocket.OPEN || ws.readyState === WebSocket.CONNECTING)) {
    return;
  }

  setConnectionStatus("connecting");
  ws = new WebSocket(WS_URL);

  ws.onopen = () => {
    reconnectAttempts = 0;
    setConnectionStatus("connected");
    startPing();
  };

  ws.onmessage = (event) => {
    handleWsMessage(event.data);
  };

  ws.onclose = (event) => {
    stopPing();
    setConnectionStatus("disconnected");
    if (!event.wasClean) {
      scheduleReconnect();
    }
  };

  ws.onerror = () => {
    // onclose will fire after this, which handles reconnection
  };
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  const delay = Math.min(
    RECONNECT_BASE_DELAY * Math.pow(2, reconnectAttempts),
    RECONNECT_MAX_DELAY
  );
  reconnectAttempts++;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectWebSocket();
  }, delay);
}

function startPing() {
  stopPing();
  pingTimer = setInterval(() => {
    if (ws && ws.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify({ type: "ping" }));
    }
  }, PING_INTERVAL);
}

function stopPing() {
  if (pingTimer) {
    clearInterval(pingTimer);
    pingTimer = null;
  }
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

function sendWsMessage(content) {
  if (!ws || ws.readyState !== WebSocket.OPEN) {
    addMessage("error", "Not connected to OpenClaw. Reconnecting...");
    connectWebSocket();
    return;
  }

  const payload = {
    type: "message",
    content: content,
    context: emailContext
      ? {
          subject: emailContext.subject,
          from: emailContext.from,
          to: emailContext.to,
          cc: emailContext.cc,
          date: emailContext.date,
          body: emailContext.body,
        }
      : null,
  };

  ws.send(JSON.stringify(payload));
}

function handleWsMessage(raw) {
  let data;
  try {
    data = JSON.parse(raw);
  } catch {
    console.warn("Non-JSON WebSocket message:", raw);
    return;
  }

  switch (data.type) {
    case "message":
    case "response":
      hideTyping();
      addMessage("ai", data.content || data.text || "");
      break;

    case "draft":
      hideTyping();
      lastDraftContent = data.content || data.text || "";
      addMessage("ai", lastDraftContent);
      $("send-reply-btn").disabled = false;
      break;

    case "error":
      hideTyping();
      addMessage("error", data.content || data.message || "An error occurred.");
      break;

    case "pong":
      // keepalive response, ignore
      break;

    default:
      if (data.content || data.text) {
        hideTyping();
        addMessage("ai", data.content || data.text);
      }
  }
}

// ===== Chat UI =====
function addMessage(role, text) {
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

function showTyping() {
  $("typing-indicator").style.display = "flex";
  scrollToBottom();
}

function hideTyping() {
  $("typing-indicator").style.display = "none";
}

function scrollToBottom() {
  const container = $("chat-messages");
  requestAnimationFrame(() => {
    container.scrollTop = container.scrollHeight;
  });
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
  sendWsMessage(text);
}

function autoResizeInput(el) {
  el.style.height = "auto";
  el.style.height = Math.min(el.scrollHeight, 100) + "px";
}

// ===== Draft Reply =====
function handleDraftReply() {
  if (!emailContext) {
    addMessage("error", "No email selected. Please open an email first.");
    return;
  }

  addMessage("user", "Draft a reply to this email");
  showTyping();

  const payload = {
    type: "draft_request",
    content: "Please draft a professional reply to this email.",
    context: emailContext
      ? {
          subject: emailContext.subject,
          from: emailContext.from,
          to: emailContext.to,
          cc: emailContext.cc,
          date: emailContext.date,
          body: emailContext.body,
        }
      : null,
  };

  if (ws && ws.readyState === WebSocket.OPEN) {
    ws.send(JSON.stringify(payload));
  } else {
    hideTyping();
    addMessage("error", "Not connected to OpenClaw. Reconnecting...");
    connectWebSocket();
  }
}

// ===== Send Reply via Office.js =====
function handleSendReply() {
  if (!lastDraftContent) {
    addMessage("error", "No draft available. Click 'Draft Reply' first.");
    return;
  }

  const item = Office.context.mailbox.item;
  if (!item) {
    addMessage("error", "No email selected.");
    return;
  }

  // Use displayReplyForm to create a reply with the draft content pre-filled.
  // This opens Outlook's native reply compose window, giving the user a chance
  // to review and edit before sending.
  try {
    item.displayReplyForm(lastDraftContent);
    addMessage("system", "Reply draft opened in Outlook. Review and send from there.");
  } catch (err) {
    console.error("Failed to create reply:", err);
    addMessage("error", "Failed to open reply form: " + err.message);
  }
}

// ===== Event Binding =====
function bindUIEvents() {
  $("send-btn").addEventListener("click", handleSend);
  $("draft-btn").addEventListener("click", handleDraftReply);
  $("send-reply-btn").addEventListener("click", handleSendReply);

  const input = $("message-input");

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  input.addEventListener("input", () => autoResizeInput(input));
}
