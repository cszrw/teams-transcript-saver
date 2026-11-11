// Runs in the page's own JavaScript world (MAIN) on Teams/SharePoint frames.
// Watches fetch/XHR for transcript metadata *and* the recap /content download,
// relays g_fileInfo / recap-link ids, and performs cookie-auth downloads.
// Also keeps 1.1.x meeting/chat observation (lib/meta.js).
(function () {
  'use strict';
  if (window.__teamsTranscriptSaverInjected) return;
  Object.defineProperty(window, '__teamsTranscriptSaverInjected', { value: true });

  const scan = window.__teamsTranscriptScan;
  const meta = window.__teamsTranscriptMeta;
  const TAG = '__teamsTranscriptSaver';
  const origFetch = window.fetch;
  let lastBearer = '';
  let lastFileInfo = null;
  let lastRecapKey = '';
  let lastItemKey = '';
  const lastDocs = [];
  let observed = 0;

  function post(msg) {
    window.postMessage(Object.assign({ [TAG]: true }, msg), location.origin);
  }

  function argUrl(args) {
    const a = args[0];
    if (typeof a === 'string') return a;
    if (typeof Request !== 'undefined' && a instanceof Request) return a.url;
    if (a && typeof a.url === 'string') return a.url;
    return '';
  }

  function argHeader(args, name) {
    const key = name.toLowerCase();
    try {
      const req = args[0];
      if (req && typeof req.headers !== 'undefined' && req.headers && typeof req.headers.get === 'function') {
        const v = req.headers.get(key);
        if (v) return v;
      }
      const h = args[1] && args[1].headers;
      if (!h) return '';
      if (typeof Headers !== 'undefined' && h instanceof Headers) return h.get(key) || '';
      if (Array.isArray(h)) {
        const e = h.find((p) => Array.isArray(p) && String(p[0]).toLowerCase() === key);
        return e ? String(e[1]) : '';
      }
      if (typeof h === 'object') {
        for (const k of Object.keys(h)) if (k.toLowerCase() === key) return h[k];
      }
    } catch (e) { /* ignore */ }
    return '';
  }

  function stamp(list) {
    const now = new Date().toISOString();
    return (list || []).map((t) => Object.assign({}, t, { capturedAt: t.capturedAt || now }));
  }

  function emit(list) {
    if (list && list.length) post({ type: 'captured', transcripts: stamp(list) });
  }

  function rememberDoc(url, json, text) {
    if (!scan || !scan.looksLikeTranscriptDoc(json) || typeof text !== 'string' || !text.trim()) return;
    lastDocs.push({ url: url || '', text, at: Date.now() });
    if (lastDocs.length > 3) lastDocs.shift();
  }

  function cachedDoc(url) {
    if (!lastDocs.length) return null;
    const match = url && lastDocs.filter((d) => d.url && d.url.split('?')[0] === String(url).split('?')[0]).pop();
    return match || lastDocs[lastDocs.length - 1];
  }

  function observe(url, status, ct, text) {
    try {
      if (!meta || observed >= 200 || typeof text !== 'string') return;
      observed++;
      const at = new Date().toISOString();
      const small = text.length < 1e6 ? text : '';
      setTimeout(() => {
        try { post({ type: 'observe', entry: Object.assign(meta.summarize(url, status, ct, small), { at, frame: location.host, bodySize: text.length }) }); }
        catch (e) { console.log('[TTS] observe failed', e); }
      }, 0);
    } catch (e) { console.log('[TTS] observe failed', e); }
  }

  function inspect(text, url) {
    try { inspectInner(text, url); } catch (e) { console.log('[TTS] inspect failed', e); }
  }

  function inspectInner(text, url) {
    const urlHint = !!(scan && url && (scan.isTranscriptishUrl(url) || scan.isSharePointItemUrl(url)));
    const hasT = !!(scan && (urlHint || scan.mightContainTranscript(text)));
    const hasM = !!(meta && text && text.length < 5e6 && meta.mightContain(text));
    if (!hasT && !hasM) return;
    let json;
    try { json = JSON.parse(text); } catch (e) {
      if (url && scan && scan.isTranscriptishUrl(url)) {
        const cap = scan.captureFromRequestUrl(url, lastFileInfo);
        if (cap) emit([cap]);
      }
      return;
    }
    if (hasM) {
      try {
        const meetings = meta.scanMeetings(json);
        const chat = meta.scanChat(json);
        if (meetings.length || chat) post({ type: 'meta', meetings, chat });
      } catch (e) { console.log('[TTS] meta parse failed', e); }
    }
    if (!scan) return;
    rememberDoc(url, json, text);
    const found = scan.scanForTranscripts(json, url);
    if (found.length) emit(found);
    else if (url && scan.isTranscriptishUrl(url)) {
      const cap = scan.captureFromRequestUrl(url, lastFileInfo);
      if (cap) emit([cap]);
    }
  }

  function rememberBearer(args, url) {
    if (!url || url.indexOf('/_api') === -1) return;
    const auth = argHeader(args, 'authorization');
    if (auth && /^Bearer\s+/i.test(auth)) lastBearer = auth;
  }

  function headersFor(url) {
    const headers = { Accept: 'application/json' };
    if (lastBearer && url && url.indexOf(location.hostname) !== -1) headers.Authorization = lastBearer;
    return headers;
  }

  async function discoverFromParts(parts, fileInfo) {
    if (!scan || !parts || !parts.driveId || !parts.itemId) return;
    const key = `${parts.driveId}/${parts.itemId}`;
    if (key === lastItemKey) return;
    lastItemKey = key;
    const tries = [
      scan.buildTranscriptUrl(parts, 'expand'),
      scan.buildTranscriptUrl(parts, 'list'),
      scan.buildTranscriptUrl(parts, 'listCurrent'),
    ].filter(Boolean);
    // Cross-origin from the Teams shell will fail; the SharePoint iframe or the
    // service worker (item-context) handles those.
    for (const url of tries) {
      try {
        const res = await origFetch.call(window, url, { credentials: 'include', cache: 'no-store', headers: headersFor(url) });
        if (!res.ok) continue;
        const text = await res.text();
        inspect(text, url);
        let json;
        try { json = JSON.parse(text); } catch (e) { json = null; }
        if (json && scan.scanForTranscripts(json, url).length) return;
      } catch (e) { /* CORS or network; try the next */ }
    }
    const cap = scan.captureFromRequestUrl(
      scan.buildTranscriptUrl(parts, 'list') || '',
      fileInfo
    );
    if (cap) emit([Object.assign(cap, fileInfo && fileInfo.name ? { itemName: fileInfo.name, name: fileInfo.displayName || fileInfo.name } : {})]);
  }

  function fileInfoParts() {
    const g = window.g_fileInfo;
    if (!g) return null;
    const sp = g['.spItemUrl'] || g.spItemUrl;
    if (!sp || !scan) return null;
    const p = scan.parseSharePointApiUrl(sp);
    if (!p) return null;
    lastFileInfo = { displayName: g.displayName, name: g.name, hasTranscripts: g.hasTranscripts };
    return p;
  }

  function watchFileInfo() {
    let tries = 0;
    const tick = () => {
      const p = fileInfoParts();
      if (p) discoverFromParts(p, lastFileInfo);
      if (++tries >= 60) clearInterval(iv);
    };
    tick();
    const iv = setInterval(tick, 1000);
    window.addEventListener('load', tick);
  }

  function watchRecap() {
    if (!scan) return;
    const tick = () => {
      const recap = scan.parseRecapHref(location.href);
      if (!recap || !recap.driveId) return;
      const key = `${recap.driveId}/${recap.itemId}`;
      if (key === lastRecapKey) return;
      lastRecapKey = key;
      if (recap.contentUrl || recap.host) {
        emit([{
          name: recap.name || 'transcript',
          url: recap.contentUrl || scan.buildTranscriptUrl(recap, 'list'),
          contentUrl: recap.contentUrl,
          driveId: recap.driveId,
          itemId: recap.itemId,
          transcriptId: recap.transcriptId,
          host: recap.host,
          origin: recap.origin,
          sitePath: recap.sitePath,
          itemName: recap.name,
        }]);
      }
      post({ type: 'item-context', item: recap });
      discoverFromParts(recap, { name: recap.name });
    };
    tick();
    window.addEventListener('hashchange', tick);
    setInterval(tick, 2000);
  }

  // --- fetch ---
  if (typeof origFetch === 'function') {
    window.fetch = function (...args) {
      const url = argUrl(args);
      rememberBearer(args, url);
      const p = origFetch.apply(this, args);
      p.then((res) => {
        try {
          const transcriptish = scan && (scan.isTranscriptishUrl(url) || scan.isSharePointItemUrl(url));
          const ct = (res.headers.get('content-type') || '');
          if (transcriptish && /cdnmedia/i.test(url)) {
            const cap = scan.captureFromRequestUrl(url, lastFileInfo);
            if (cap) emit([cap]);
            post({ type: 'item-context', item: scan.parseSharePointApiUrl(url) });
            return;
          }
          if (transcriptish || /json/i.test(ct)) {
            res.clone().text().then((text) => {
              if (/json/i.test(ct)) observe(res.url || url, res.status, ct, text);
              inspect(text, url || res.url);
            }, () => {});
          }
        } catch (e) { /* never break the page */ }
      }, () => {});
      return p;
    };
  }

  // --- XMLHttpRequest ---
  const XHR = window.XMLHttpRequest && window.XMLHttpRequest.prototype;
  if (XHR) {
    const origOpen = XHR.open;
    const origSend = XHR.send;
    XHR.open = function (method, url, ...rest) {
      this.__ttsUrl = typeof url === 'string' ? url : '';
      return origOpen.call(this, method, url, ...rest);
    };
    XHR.send = function (...args) {
      this.addEventListener('load', function () {
        try {
          const url = this.__ttsUrl || this.responseURL || '';
          const ct = this.getResponseHeader('content-type') || '';
          const transcriptish = scan && (scan.isTranscriptishUrl(url) || scan.isSharePointItemUrl(url));
          if (transcriptish && /cdnmedia/i.test(url)) {
            const cap = scan.captureFromRequestUrl(url, lastFileInfo);
            if (cap) emit([cap]);
            post({ type: 'item-context', item: scan.parseSharePointApiUrl(url) });
            return;
          }
          if (!(transcriptish || /json/i.test(ct))) return;
          let t = null;
          if (this.responseType === '' || this.responseType === 'text') t = this.responseText;
          else if (this.responseType === 'json' && this.response) t = JSON.stringify(this.response);
          if (t != null) {
            if (/json/i.test(ct)) observe(this.responseURL || url, this.status, ct, t);
            inspect(t, url);
          }
        } catch (e) { /* never break the page */ }
      });
      return origSend.apply(this, args);
    };
  }

  // --- download on request from content.js ---
  async function download(url) {
    let lastErr;
    let lastRes;
    const headerSets = lastBearer ? [{ Authorization: lastBearer }, {}] : [{}];
    for (const headers of headerSets) {
      for (const credentials of ['include', 'omit']) {
        try {
          const res = await origFetch.call(window, url, { credentials, cache: 'no-store', headers });
          const text = res.ok ? await res.text() : '';
          lastRes = { ok: res.ok, status: res.status, text };
          if (res.ok && text && text.trim()) return lastRes;
        } catch (e) {
          lastErr = e;
        }
      }
    }
    // If SharePoint /content 403s (Download button disabled), reuse a transcript
    // JSON body this frame already loaded — that does not need recording-download rights.
    const cached = cachedDoc(url);
    if (cached) return { ok: true, status: 200, text: cached.text };
    if (lastRes) return lastRes;
    return { ok: false, status: 0, error: String(lastErr && lastErr.message || lastErr) };
  }

  window.addEventListener('message', (e) => {
    const d = e.data;
    if (e.source !== window || !d || d[TAG] !== true || d.type !== 'fetch-request') return;
    download(d.url).then((r) => post(Object.assign({ type: 'fetch-response', id: d.id }, r)));
  });

  watchFileInfo();
  watchRecap();
})();
