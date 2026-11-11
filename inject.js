// Runs in the page's own JavaScript world (MAIN) on Teams/SharePoint frames.
// Watches fetch/XHR JSON responses for transcript download links and hands them
// to content.js. Also performs the transcript download in the page context when asked.
(function () {
  'use strict';
  if (window.__teamsTranscriptSaverInjected) return;
  Object.defineProperty(window, '__teamsTranscriptSaverInjected', { value: true });

  const scan = window.__teamsTranscriptScan;
  const TAG = '__teamsTranscriptSaver';
  const origFetch = window.fetch;

  function post(msg) {
    window.postMessage(Object.assign({ [TAG]: true }, msg), location.origin);
  }

  const meta = window.__teamsTranscriptMeta;

  let observed = 0;
  function observe(url, status, ct, text) {
    try {
      if (!meta || observed >= 200 || typeof text !== 'string') return;
      observed++;
      const at = new Date().toISOString();
      const small = text.length < 1e6 ? text : '';
      // Lazy: summarise off the response path, and send only the small summary.
      setTimeout(() => {
        try { post({ type: 'observe', entry: Object.assign(meta.summarize(url, status, ct, small), { at, frame: location.host, bodySize: text.length }) }); }
        catch (e) { console.log('[TTS] observe failed', e); }
      }, 0);
    } catch (e) { console.log('[TTS] observe failed', e); }
  }
  function inspect(text) { try { inspectInner(text); } catch (e) { console.log('[TTS] inspect failed', e); } }
  function inspectInner(text) {
    const hasT = scan.mightContainTranscript(text);
    const hasM = meta && text && text.length < 5e6 && meta.mightContain(text);
    if (!hasT && !hasM) return;
    let json;
    try { json = JSON.parse(text); } catch (e) { return; }
    if (hasM) {
      try {
        const meetings = meta.scanMeetings(json);
        const chat = meta.scanChat(json);
        if (meetings.length || chat) post({ type: 'meta', meetings, chat });
      } catch (e) { console.log('[TTS] meta parse failed', e); }
    }
    if (!hasT) return;
    const found = scan.scanForTranscripts(json);
    if (found.length) post({ type: 'captured', transcripts: found.map((t) => Object.assign(t, { capturedAt: new Date().toISOString() })) });
  }

  // --- fetch ---
  if (typeof origFetch === 'function') {
    window.fetch = function (...args) {
      const p = origFetch.apply(this, args);
      p.then((res) => {
        try {
          const ct = res.headers.get('content-type') || '';
          if (!/json/i.test(ct)) return;
          res.clone().text().then((t) => { observe(res.url, res.status, ct, t); inspect(t); }, () => {});
        } catch (e) { /* never break the page */ }
      }, () => {});
      return p;
    };
  }

  // --- XMLHttpRequest ---
  const XHR = window.XMLHttpRequest && window.XMLHttpRequest.prototype;
  if (XHR) {
    const origSend = XHR.send;
    XHR.send = function (...args) {
      this.addEventListener('load', function () {
        try {
          const ct = this.getResponseHeader('content-type') || '';
          if (!/json/i.test(ct)) return;
          let t = null;
          if (this.responseType === '' || this.responseType === 'text') t = this.responseText;
          else if (this.responseType === 'json' && this.response) t = JSON.stringify(this.response);
          if (t != null) { observe(this.responseURL, this.status, ct, t); inspect(t); }
        } catch (e) { /* never break the page */ }
      });
      return origSend.apply(this, args);
    };
  }

  // --- download on request from content.js ---
  async function download(url) {
    let lastErr;
    for (const credentials of ['include', 'omit']) {
      try {
        const res = await origFetch.call(window, url, { credentials, cache: 'no-store' });
        return { ok: res.ok, status: res.status, text: res.ok ? await res.text() : '' };
      } catch (e) {
        lastErr = e; // usually CORS; try the next mode
      }
    }
    return { ok: false, status: 0, error: String(lastErr && lastErr.message || lastErr) };
  }

  window.addEventListener('message', (e) => {
    const d = e.data;
    if (e.source !== window || !d || d[TAG] !== true || d.type !== 'fetch-request') return;
    download(d.url).then((r) => post(Object.assign({ type: 'fetch-response', id: d.id }, r)));
  });
})();
