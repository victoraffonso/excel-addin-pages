/**
 * Bridge Client — runs inside the Office.js task pane.
 * Connects to the MCP server's WebSocket bridge and executes Excel operations.
 */
import { executeOperation } from './operations.js';

const WS_URL = 'wss://localhost:3100';
const RECONNECT_DELAY = 2000;

let ws = null;
let statusEl = null;
let logEl = null;
let statsEl = null;
let opCount = 0;
let errCount = 0;

function log(msg) {
  if (!logEl) return;
  const entry = document.createElement('div');
  entry.className = 'log-entry';
  entry.textContent = `${new Date().toLocaleTimeString()} ${msg}`;
  logEl.prepend(entry);
  // Keep max 100 entries
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

function connect() {
  setStatus('connecting', 'Connecting...');
  log('Connecting to bridge...');

  ws = new WebSocket(WS_URL);

  ws.onopen = () => {
    setStatus('connected', 'Connected to MCP Bridge');
    log('Connected');
  };

  ws.onclose = () => {
    setStatus('disconnected', 'Disconnected — reconnecting...');
    log('Disconnected');
    ws = null;
    setTimeout(connect, RECONNECT_DELAY);
  };

  ws.onerror = (err) => {
    log(`WebSocket error: ${err.type}`);
    // onclose will handle reconnection
  };

  ws.onmessage = async (event) => {
    let msg;
    try {
      msg = JSON.parse(event.data);
    } catch (e) {
      log(`Bad message: ${e}`);
      return;
    }

    if (msg.batch) {
      // Batch execution — multiple operations in single Excel.run
      await handleBatch(msg);
    } else {
      // Single operation
      await handleSingle(msg);
    }
  };
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

    ws?.send(JSON.stringify({ id, ok: true, result }));
  } catch (error) {
    errCount++;
    const ms = (performance.now() - start).toFixed(0);
    log(`${tool} ERROR (${ms}ms): ${error.message}`);
    updateStats();

    ws?.send(JSON.stringify({ id, ok: false, error: error.message }));
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
      await context.sync(); // Single sync for all operations
      return res;
    });

    opCount += operations.length;
    const ms = (performance.now() - start).toFixed(0);
    log(`batch(${operations.length} ops) → ${ms}ms`);
    updateStats();

    ws?.send(JSON.stringify({ id, ok: true, result: results }));
  } catch (error) {
    errCount++;
    const ms = (performance.now() - start).toFixed(0);
    log(`batch ERROR (${ms}ms): ${error.message}`);
    updateStats();

    ws?.send(JSON.stringify({ id, ok: false, error: error.message }));
  }
}

export function initBridge(statusElement, logElement, statsElement) {
  statusEl = statusElement;
  logEl = logElement;
  statsEl = statsElement;
  connect();
}
