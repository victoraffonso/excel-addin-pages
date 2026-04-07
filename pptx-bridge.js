/**
 * Bridge Client — runs inside the Office.js task pane for PowerPoint.
 * Polls the MCP server's HTTPS bridge for pending commands and executes them via PowerPoint.run().
 * Uses HTTP polling instead of WebSocket (WebSocket blocked by Office WebView sandbox).
 */
import { executeOperation } from './pptx-ops.js';

const BRIDGE_URL = 'https://localhost:3100';
const IDLE_POLL = 500;   // ms when no pending commands
const ACTIVE_POLL = 100; // ms when commands are being processed
let currentPollInterval = IDLE_POLL;

let statusEl = null;
let logEl = null;
let statsEl = null;
let opCount = 0;
let errCount = 0;
let polling = false;
let presentationId = ''; // Identifies this presentation to the bridge
let addinToken = ''; // Auth token from bridge

function log(msg) {
  if (!logEl) return;
  const entry = document.createElement('div');
  entry.className = 'log-entry';
  entry.textContent = `${new Date().toLocaleTimeString()} ${msg}`;
  logEl.prepend(entry);
  while (logEl.children.length > 100) logEl.removeChild(logEl.lastChild);
}

function setStatus(state, text) {
  if (!statusEl) return;
  statusEl.className = `status ${state}`;
  statusEl.textContent = text;
}

function updateStats() {
  if (statsEl) statsEl.textContent = `Ops: ${opCount} | Errors: ${errCount}`;
}

async function resolvePresentationId() {
  try {
    // Office.context.document.url gives the file path/URL of the active document
    const url = Office.context.document.url || '';
    // Extract filename from path or URL
    const parts = url.replace(/\\/g, '/').split('/');
    presentationId = parts[parts.length - 1] || 'Untitled';
    log(`Presentation: ${presentationId}`);
  } catch (e) {
    log(`Failed to get presentation name: ${e.message}`);
    presentationId = 'Unknown';
  }
}

async function poll() {
  if (!polling) return;

  // Resolve presentation ID on first poll
  if (!presentationId) {
    await resolvePresentationId();
  }

  try {
    // Get auth token on first poll
    if (!addinToken) {
      try {
        const tokenResp = await fetch(`${BRIDGE_URL}/addin-token`);
        if (tokenResp.ok) {
          const data = await tokenResp.json();
          addinToken = data.token || '';
          log('Auth token obtained');
        }
      } catch {}
    }

    const resp = await fetch(`${BRIDGE_URL}/poll?workbook=${encodeURIComponent(presentationId)}&token=${addinToken}`, { method: 'GET' });

    if (resp.status === 204) {
      setStatus('connected', `Connected — ${presentationId}`);
      currentPollInterval = IDLE_POLL;
      setTimeout(poll, currentPollInterval);
      return;
    }

    if (resp.status === 200) {
      setStatus('connected', `Connected — ${presentationId}`);
      currentPollInterval = ACTIVE_POLL; // Speed up polling when processing commands
      const msg = await resp.json();

      if (msg.batch) {
        await handleBatch(msg);
      } else {
        await handleSingle(msg);
      }
    }
  } catch (e) {
    setStatus('disconnected', 'Disconnected — retrying...');
    // Server not reachable — retry after delay
  }

  setTimeout(poll, currentPollInterval);
}

async function handleSingle(msg) {
  const { id, tool, args } = msg;
  const start = performance.now();

  try {
    const result = await PowerPoint.run(async (context) => {
      const res = await executeOperation(context, tool, args);
      await context.sync();
      return res;
    });

    opCount++;
    const ms = (performance.now() - start).toFixed(0);
    log(`${tool} → ${ms}ms`);
    updateStats();

    await sendResult({ id, ok: true, result });
  } catch (error) {
    errCount++;
    const ms = (performance.now() - start).toFixed(0);
    log(`${tool} ERROR (${ms}ms): ${error.message}`);
    updateStats();

    await sendResult({ id, ok: false, error: error.message });
  }
}

async function handleBatch(msg) {
  const { id, operations } = msg;
  const start = performance.now();

  try {
    const results = await PowerPoint.run(async (context) => {
      const res = [];
      for (const op of operations) {
        res.push(await executeOperation(context, op.tool, op.args));
      }
      await context.sync();
      return res;
    });

    opCount += operations.length;
    const ms = (performance.now() - start).toFixed(0);
    log(`batch(${operations.length} ops) → ${ms}ms`);
    updateStats();

    await sendResult({ id, ok: true, result: results });
  } catch (error) {
    errCount++;
    const ms = (performance.now() - start).toFixed(0);
    log(`batch ERROR (${ms}ms): ${error.message}`);
    updateStats();

    await sendResult({ id, ok: false, error: error.message });
  }
}

async function sendResult(data) {
  try {
    await fetch(`${BRIDGE_URL}/result?token=${addinToken}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
  } catch (e) {
    log(`Failed to send result: ${e.message}`);
  }
}

export function initBridge(statusElement, logElement, statsElement) {
  statusEl = statusElement;
  logEl = logElement;
  statsEl = statsElement;

  setStatus('connecting', 'Connecting...');
  log('Starting HTTP polling (PowerPoint)...');
  polling = true;
  poll();
}
