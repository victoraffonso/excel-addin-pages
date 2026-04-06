/**
 * Bridge Client — runs inside the Office.js task pane.
 * Polls the MCP server's HTTPS bridge for pending commands and executes them via Excel.run().
 * Uses HTTP polling instead of WebSocket (WebSocket blocked by Excel's WebView sandbox).
 */
import { executeOperation } from './operations.js';

const BRIDGE_URL = 'https://localhost:3100';
const POLL_INTERVAL = 150; // ms between polls

let statusEl = null;
let logEl = null;
let statsEl = null;
let opCount = 0;
let errCount = 0;
let polling = false;

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

async function poll() {
  if (!polling) return;

  try {
    const resp = await fetch(`${BRIDGE_URL}/poll`, { method: 'GET' });

    if (resp.status === 204) {
      // No pending commands — poll again
      setTimeout(poll, POLL_INTERVAL);
      return;
    }

    if (resp.status === 200) {
      setStatus('connected', 'Connected to MCP Bridge');
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

  setTimeout(poll, POLL_INTERVAL);
}

async function handleSingle(msg) {
  const { id, tool, args } = msg;
  const start = performance.now();

  try {
    const result = await Excel.run(async (context) => {
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
    const results = await Excel.run(async (context) => {
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
    await fetch(`${BRIDGE_URL}/result`, {
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
  log('Starting HTTP polling...');
  polling = true;
  poll();
}
