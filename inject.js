// Runs in the page's own JavaScript world (MAIN) on Teams/SharePoint frames.
// Watches fetch/XHR for transcript metadata *and* the recap /content download,
// relays g_fileInfo / recap-link ids, and performs cookie-auth downloads.
(function () {
  'use strict';
  if (window.__teamsTranscriptSaverInjected) return;
  Object.defineProperty(window, '__teamsTranscriptSaverInjected', { value: true });

  const scan = window.__teamsTranscriptScan;
  const TAG = '__teamsTranscriptSaver';
  const origFetch = window.fetch;
  let lastBearer = '';
  let lastFileInfo = null;
  let lastRecapKey = '';
  let lastItemKey = '';

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

  function inspect(text, url) {
    if (!scan) return;
    const urlHint = url && (scan.isTranscriptishUrl(url) || scan.isSharePointItemUrl(url));
    if (!urlHint && !scan.mightContainTranscript(text)) return;
    let json;
    try { json = JSON.parse(text); } catch (e) {
      if (url && scan.isTranscriptishUrl(url)) {
        const cap = scan.captureFromRequestUrl(url, lastFileInfo);
        if (cap) emit([cap]);
      }
      return;
    }
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
            res.clone().text().then((text) => inspect(text, url), () => {});
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
          const url = this.__ttsUrl || '';
          const ct = this.getResponseHeader('content-type') || '';
          const transcriptish = scan && (scan.isTranscriptishUrl(url) || scan.isSharePointItemUrl(url));
          if (transcriptish && /cdnmedia/i.test(url)) {
            const cap = scan.captureFromRequestUrl(url, lastFileInfo);
            if (cap) emit([cap]);
            return;
          }
          if (!(transcriptish || /json/i.test(ct))) return;
          if (this.responseType === '' || this.responseType === 'text') inspect(this.responseText, url);
          else if (this.responseType === 'json' && this.response) inspect(JSON.stringify(this.response), url);
        } catch (e) { /* never break the page */ }
      });
      return origSend.apply(this, args);
    };
  }

  // --- download on request from content.js ---
  async function download(url) {
    let lastErr;
    const headerSets = lastBearer ? [{ Authorization: lastBearer }, {}] : [{}];
    for (const headers of headerSets) {
      for (const credentials of ['include', 'omit']) {
        try {
          const res = await origFetch.call(window, url, { credentials, cache: 'no-store', headers });
          return { ok: res.ok, status: res.status, text: res.ok ? await res.text() : '' };
        } catch (e) {
          lastErr = e;
        }
      }
    }
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
