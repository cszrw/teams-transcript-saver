// Isolated-world content script. Bridges inject.js <-> background.js and,
// in the top frame only, shows the floating "Save transcript" button.
(function () {
  'use strict';
  const TAG = '__teamsTranscriptSaver';
  const isTop = window === window.top;
  const pending = new Map();
  let nextId = 1;

  // Page -> background
  window.addEventListener('message', (e) => {
    const d = e.data;
    if (e.source !== window || !d || d[TAG] !== true) return;
    if (d.type === 'captured') {
      chrome.runtime.sendMessage({ type: 'captured', transcripts: d.transcripts }).catch(() => {});
    } else if (d.type === 'fetch-response' && pending.has(d.id)) {
      pending.get(d.id)(d);
      pending.delete(d.id);
    }
  });

  // Ask inject.js (page context) to download a URL.
  function pageFetch(url) {
    return new Promise((resolve) => {
      const id = nextId++;
      pending.set(id, resolve);
      setTimeout(() => { if (pending.delete(id)) resolve({ ok: false, status: 0, error: 'timeout' }); }, 30000);
      window.postMessage({ [TAG]: true, type: 'fetch-request', id, url }, location.origin);
    });
  }

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === 'page-fetch') {
      pageFetch(msg.url).then(sendResponse);
      return true;
    }
    if (msg.type === 'state' && isTop) renderButton(msg.count);
    if (msg.type === 'status' && isTop) showStatus(msg.text, msg.error);
    return false;
  });

  if (!isTop) return;

  // ---- floating button (top frame only) ----
  let host, btn, note;

  function ensureUi() {
    if (host) return;
    host = document.createElement('div');
    host.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;';
    const shadow = host.attachShadow({ mode: 'closed' });
    shadow.innerHTML = `
      <style>
        .wrap{font:13px/1.3 -apple-system,"Segoe UI",sans-serif;display:flex;flex-direction:column;align-items:flex-end;gap:6px}
        button{all:unset;cursor:pointer;background:#5b5fc7;color:#fff;padding:7px 12px;border-radius:16px;
          box-shadow:0 2px 8px rgba(0,0,0,.25);opacity:.85;user-select:none}
        button:hover{opacity:1}
        .note{background:#fff;color:#222;border:1px solid #ccc;border-radius:6px;padding:6px 9px;max-width:280px;
          box-shadow:0 2px 8px rgba(0,0,0,.2);display:none}
        .note.err{border-color:#c4314b;color:#a4262c}
      </style>
      <div class="wrap"><div class="note"></div><button title="Drag to move">⬇ Save transcript</button></div>`;
    btn = shadow.querySelector('button');
    note = shadow.querySelector('.note');
    makeDraggable();
    (document.body || document.documentElement).appendChild(host);
  }

  function makeDraggable() {
    let start = null;
    let moved = false;
    btn.addEventListener('pointerdown', (e) => {
      const r = host.getBoundingClientRect();
      start = { x: e.clientX, y: e.clientY, right: innerWidth - r.right, bottom: innerHeight - r.bottom };
      moved = false;
      btn.setPointerCapture(e.pointerId);
    });
    btn.addEventListener('pointermove', (e) => {
      if (!start) return;
      const dx = e.clientX - start.x, dy = e.clientY - start.y;
      if (Math.abs(dx) + Math.abs(dy) > 4) moved = true;
      if (!moved) return;
      host.style.right = Math.max(0, Math.min(innerWidth - 60, start.right - dx)) + 'px';
      host.style.bottom = Math.max(0, Math.min(innerHeight - 30, start.bottom - dy)) + 'px';
    });
    btn.addEventListener('pointerup', () => {
      const wasClick = start && !moved;
      start = null;
      if (wasClick) save();
    });
  }

  function save() {
    showStatus('Saving…');
    chrome.runtime.sendMessage({ type: 'save-latest' }).catch(() => showStatus('Extension was reloaded. Reload this page.', true));
  }

  function renderButton(count) {
    if (!count) { if (host) host.style.display = 'none'; return; }
    ensureUi();
    host.style.display = '';
    btn.textContent = count > 1 ? `⬇ Save transcript (${count})` : '⬇ Save transcript';
    btn.title = count > 1 ? 'Saves the most recent one. Use the toolbar icon to pick another. Drag to move.' : 'Drag to move';
  }

  let noteTimer;
  function showStatus(text, isError) {
    if (!host) return;
    note.textContent = text;
    note.className = 'note' + (isError ? ' err' : '');
    note.style.display = 'block';
    clearTimeout(noteTimer);
    noteTimer = setTimeout(() => { note.style.display = 'none'; }, isError ? 12000 : 5000);
  }

  // Ask for current state (e.g. after an in-page reload of the content script).
  chrome.runtime.sendMessage({ type: 'hello' }).then((r) => r && renderButton(r.count)).catch(() => {});
})();
