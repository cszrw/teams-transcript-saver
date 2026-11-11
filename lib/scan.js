// Finds transcript download links in a JSON API response, and turns Teams/SharePoint
// request URLs into cookie-authenticated download URLs.
// Shared by inject.js (page), background.js, and the node tests. No dependencies.
(function (root) {
  'use strict';

  const MAX_DEPTH = 12;
  const DOWNLOAD_KEYS = [
    'temporaryDownloadUrl',
    'downloadUrl',
    '@microsoft.graph.downloadUrl',
    '@content.downloadUrl',
    'contentDownloadUrl',
  ];

  function str(v) {
    return typeof v === 'string' && v ? v : undefined;
  }

  function ticksToIso(ticks) {
    const n = typeof ticks === 'number' ? ticks : (typeof ticks === 'string' && /^\d+$/.test(ticks) ? +ticks : NaN);
    if (!isFinite(n) || n < 6e16) return undefined;
    const ms = (n - 621355968000000000) / 10000;
    const d = new Date(ms);
    return isNaN(d.getTime()) ? undefined : d.toISOString();
  }

  function downloadUrlOf(obj) {
    if (!obj || typeof obj !== 'object') return undefined;
    for (const k of DOWNLOAD_KEYS) {
      const v = obj[k];
      if (typeof v === 'string' && /^https:\/\//i.test(v)) return v;
    }
    return undefined;
  }

  function hostnameOf(url) {
    try { return new URL(url).hostname; } catch (e) { return ''; }
  }

  // Parse a SharePoint drive/item/transcript URL, including the Teams recap
  // `_api_cached/.../cdnmedia/transcripts` form (binary; IDs only).
  function parseSharePointApiUrl(url) {
    if (typeof url !== 'string' || url.indexOf('/_api') === -1) return null;
    let u;
    try { u = new URL(url); } catch (e) { return null; }
    const m = u.pathname.match(
      /^(.*)\/_api(?:_cached)?\/v[0-9.]+\/drives\/([^/]+)\/items\/([^/]+)(?:\/versions\/[^/]+)?(?:\/(?:cdnmedia|media)\/transcripts(?:\/([^/]+))?(?:\/(content|streamContent))?)?/i
    );
    if (!m) return null;
    const sitePath = (m[1] || '').replace(/\/$/, '');
    const kind = m[5] ? m[5].toLowerCase()
      : /\/(?:cdnmedia|media)\/transcripts/i.test(u.pathname) ? 'list'
        : 'item';
    return {
      origin: u.origin,
      host: u.hostname,
      sitePath,
      driveId: decodeURIComponent(m[2]),
      itemId: decodeURIComponent(m[3]),
      transcriptId: m[4] ? decodeURIComponent(m[4]) : undefined,
      kind,
    };
  }

  function partsFromCapture(t) {
    if (!t) return null;
    const fromUrl = parseSharePointApiUrl(t.contentUrl || t.url || t.location || '');
    if (fromUrl) {
      return {
        origin: fromUrl.origin,
        host: fromUrl.host,
        sitePath: fromUrl.sitePath || t.sitePath || '',
        driveId: t.driveId || fromUrl.driveId,
        itemId: t.itemId || fromUrl.itemId,
        transcriptId: t.transcriptId || fromUrl.transcriptId,
      };
    }
    if (t.driveId && t.itemId && (t.origin || t.host)) {
      const host = t.host || hostnameOf(t.origin);
      const origin = t.origin || (host ? 'https://' + host : '');
      if (!origin) return null;
      return {
        origin,
        host,
        sitePath: t.sitePath || '',
        driveId: t.driveId,
        itemId: t.itemId,
        transcriptId: t.transcriptId,
      };
    }
    return null;
  }

  // Cookie-authenticated SharePoint REST URLs. Always `_api/v2.1` + `/media/`,
  // never `_api_cached` or `/cdnmedia/` (those need a CDN token and 401 today).
  function buildTranscriptUrl(p, which, format) {
    if (!p || !p.origin || !p.driveId || !p.itemId) return null;
    const site = p.sitePath || '';
    const base = `${p.origin}${site}/_api/v2.1/drives/${p.driveId}/items/${p.itemId}`;
    let url;
    if (which === 'expand') {
      url = `${base}?select=${encodeURIComponent('media/transcripts')}&$expand=${encodeURIComponent('media/transcripts')}`;
    } else if (which === 'list') {
      url = `${base}/media/transcripts`;
    } else if (which === 'listCurrent') {
      url = `${base}/versions/current/media/transcripts`;
    } else if (which === 'metadata' && p.transcriptId) {
      url = `${base}/media/transcripts/${p.transcriptId}`;
    } else if ((which === 'content' || which === 'streamContent') && p.transcriptId) {
      url = `${base}/media/transcripts/${p.transcriptId}/${which}`;
    } else {
      return null;
    }
    if (format && (which === 'content' || which === 'streamContent')) {
      url += (url.indexOf('?') === -1 ? '?' : '&') + 'format=' + encodeURIComponent(format);
    }
    return url;
  }

  function withFormat(url, format) {
    if (!url) return url;
    try {
      const u = new URL(url);
      u.searchParams.set('format', format);
      return u.toString();
    } catch (e) { return url; }
  }

  // Stream often wants these query flags; they do not require recording-download rights.
  function withEditsOff(url) {
    if (!url) return url;
    try {
      const u = new URL(url);
      u.searchParams.set('applyhighlights', 'false');
      u.searchParams.set('applymediaedits', 'false');
      return u.toString();
    } catch (e) { return url; }
  }

  function urlKind(url) {
    if (!url) return 'none';
    if (/cdnmedia/i.test(url)) return 'cdnmedia';
    if (/[?&]tempauth=/i.test(url)) return 'tempauth';
    if (/\/content(?:\?|$)/i.test(url)) return 'content';
    if (/\/streamContent/i.test(url)) return 'streamContent';
    if (/\/media\/transcripts/i.test(url)) return 'media-transcripts';
    if (isCollabUrl(url)) return 'collab';
    return 'other';
  }

  function isCdnOrTempauth(url) {
    const k = urlKind(url);
    return k === 'cdnmedia' || k === 'tempauth';
  }

  function isCollabUrl(url) {
    if (typeof url !== 'string' || !/^https:\/\//i.test(url)) return false;
    try {
      const u = new URL(url);
      if (!/(^|\.)(teams\.cloud\.microsoft|teams\.microsoft\.com)$/i.test(u.hostname)) return false;
      return /\/api\/mcps\//i.test(u.pathname) || /collab/i.test(u.pathname + u.search);
    } catch (e) { return false; }
  }

  // Recap links: teams.cloud.microsoft/l/meetingrecap?... and the `/_#/l/` form.
  function parseRecapHref(href) {
    if (typeof href !== 'string') return null;
    let u;
    try { u = new URL(href); } catch (e) { return null; }
    const hash = u.hash || '';
    const hashQuery = hash.includes('?') ? hash.slice(hash.indexOf('?') + 1) : '';
    const q = new URLSearchParams(u.search);
    new URLSearchParams(hashQuery).forEach((v, k) => { if (!q.has(k)) q.set(k, v); });
    const pathish = `${u.pathname}${hash}`.toLowerCase();
    if (!/meetingrecap/.test(pathish) && !q.get('driveId')) return null;
    const driveId = str(q.get('driveId'));
    const itemId = str(q.get('driveItemId')) || str(q.get('itemId'));
    if (!driveId || !itemId) return null;
    const siteRaw = q.get('sitePath') || '';
    const fileUrl = q.get('fileUrl') || q.get('webUrl') || '';
    let host, origin, sitePath = '';
    const hint = siteRaw || fileUrl;
    if (hint) {
      try {
        if (/^https?:/i.test(hint)) {
          const s = new URL(hint);
          host = s.hostname;
          origin = s.origin;
          sitePath = s.pathname
            .replace(/\/_api(?:_cached)?\/.*$/i, '')
            .replace(/\/_layouts\/.*$/i, '')
            .replace(/\/Documents\/.*$/i, '')
            .replace(/\/+$/, '');
        } else if (siteRaw.charAt(0) === '/') {
          sitePath = siteRaw.replace(/\/+$/, '');
        }
      } catch (e) { /* ignore */ }
    }
    const parsed = parseSharePointApiUrl(siteRaw) || parseSharePointApiUrl(fileUrl);
    if (parsed) {
      host = host || parsed.host;
      origin = origin || parsed.origin;
      sitePath = sitePath || parsed.sitePath;
    }
    let name;
    if (fileUrl) {
      try {
        name = decodeURIComponent(new URL(fileUrl).pathname.split('/').pop() || '');
      } catch (e) { /* ignore */ }
    }
    const transcriptId = parsed && parsed.transcriptId;
    const parts = host ? { origin: origin || ('https://' + host), host, sitePath, driveId, itemId, transcriptId } : null;
    return {
      driveId,
      itemId,
      transcriptId,
      host,
      origin: parts && parts.origin,
      sitePath,
      name,
      contentUrl: parts && transcriptId ? buildTranscriptUrl(parts, 'content') : undefined,
    };
  }

  function meetingDateOf(node, json) {
    const media = json && json.media;
    return str(node.meetingStartDateTime)
      || str(node.recordingStartDateTime)
      || ticksToIso(node.startTime)
      || (node.metadata && (str(node.metadata.meetingStartDateTime) || ticksToIso(node.metadata.startTime)))
      || str(json && json.meetingStartDateTime)
      || (media && (str(media.meetingStartDateTime) || str(media.recordingStartDateTime) || ticksToIso(media.startTime)));
  }

  function looksLikeTranscriptDoc(json) {
    if (!json || typeof json !== 'object' || Array.isArray(json)) return false;
    if (json.$schema && /transcript/i.test(String(json.$schema))) return true;
    if (String(json.type || '') === 'Transcript' && Array.isArray(json.entries)) return true;
    const e = Array.isArray(json.entries) ? json.entries[0] : null;
    return !!(e && typeof e === 'object' && (e.text || e.content) && (e.startOffset != null || e.start != null));
  }

  function looksLikeMcpsTranscript(obj) {
    if (!obj || typeof obj !== 'object') return false;
    const type = str(obj.type) || '';
    const loc = str(obj.location) || '';
    const md = obj.metadata && typeof obj.metadata === 'object' ? obj.metadata : null;
    if (/transcript/i.test(type)) return true;
    if (/\/(?:media|cdnmedia)\/transcripts\//i.test(loc)) return true;
    if (md && str(md.transcriptId) && (str(md.driveId) || str(md.driveItemId))) return true;
    return false;
  }

  function isTranscriptMetadataRow(obj, path, fromSrc) {
    if (!obj || typeof obj !== 'object') return false;
    if (!(obj.transcriptType || obj.languageTag || obj.displayName || obj.isDefault != null)) return false;
    if (path.some((k) => /transcript/i.test(k))) return true;
    if (fromSrc && (fromSrc.kind === 'list' || fromSrc.kind === 'content' || fromSrc.kind === 'streamcontent')) return true;
    return false;
  }

  function looksLikeTranscript(obj, path, fromSrc) {
    if (looksLikeMcpsTranscript(obj)) return true;
    const url = downloadUrlOf(obj);
    const inTranscriptPath = path.some((k) => /transcript/i.test(k));
    if (url) {
      if (inTranscriptPath || 'transcriptType' in obj || 'languageTag' in obj) return true;
      if (/transcript/i.test(url.split('?')[0])) return true;
    }
    if (isTranscriptMetadataRow(obj, path, fromSrc) && str(obj.id) &&
        (obj.transcriptType || obj.languageTag || obj.displayName || obj.isDefault != null)) {
      return true;
    }
    return false;
  }

  function captureFrom(node, path, parentItem, rootItemName, rootCreated, json, fromSrc) {
    const md = node.metadata && typeof node.metadata === 'object' ? node.metadata : null;
    const loc = str(node.location);
    const parsedLoc = parseSharePointApiUrl(downloadUrlOf(node) || loc || '');
    const driveId = str(node.driveId) || (md && str(md.driveId)) || (parsedLoc && parsedLoc.driveId) || (fromSrc && fromSrc.driveId);
    const itemId = str(node.driveItemId) || str(node.itemId) || (md && (str(md.driveItemId) || str(md.itemId)))
      || (parsedLoc && parsedLoc.itemId) || (fromSrc && fromSrc.itemId);
    const idLooksLikeTranscript = str(node.id) && (
      isTranscriptMetadataRow(node, path, fromSrc) || looksLikeMcpsTranscript(node) || 'transcriptType' in node
    );
    const transcriptId = (idLooksLikeTranscript && str(node.id))
      || (md && str(md.transcriptId))
      || (parsedLoc && parsedLoc.transcriptId)
      || (fromSrc && fromSrc.transcriptId)
      || undefined;
    const host = (parsedLoc && parsedLoc.host) || (fromSrc && fromSrc.host)
      || (loc ? hostnameOf(loc) : undefined);
    const origin = (parsedLoc && parsedLoc.origin) || (fromSrc && fromSrc.origin)
      || (host ? 'https://' + host : undefined);
    const sitePath = (parsedLoc && parsedLoc.sitePath) || (fromSrc && fromSrc.sitePath) || '';
    const parts = (driveId && itemId && origin)
      ? { origin, host: host || hostnameOf(origin), sitePath, driveId, itemId, transcriptId }
      : null;
    const contentUrl = parts && transcriptId ? buildTranscriptUrl(parts, 'content') : undefined;
    const raw = downloadUrlOf(node);
    const url = contentUrl || raw || (loc && /^https:\/\//i.test(loc) && !isCdnOrTempauth(loc) ? loc.split('?')[0] : undefined)
      || (parts && transcriptId ? buildTranscriptUrl(parts, 'streamContent') : undefined);
    if (!url && !contentUrl && !parts) return null;
    const item = parentItem;
    return {
      name: str(node.displayName) || str(node.name) || str(node.fileName)
        || (md && str(md.fileTitle)) || (item && str(item.name)) || rootItemName || 'transcript',
      language: str(node.languageTag) || str(node.language) || str(node.locale)
        || (md && str(md.languageTag)),
      transcriptType: str(node.transcriptType),
      url: url || contentUrl,
      contentUrl,
      location: loc,
      driveId,
      itemId,
      transcriptId,
      host: parts && parts.host,
      origin: parts && parts.origin,
      sitePath: parts && parts.sitePath,
      itemName: (item && item !== node && str(item.name)) || (md && str(md.fileTitle)) || rootItemName,
      createdDateTime: str(node.createdDateTime) || (item && str(item.createdDateTime)) || rootCreated,
      meetingDate: meetingDateOf(node, json) || meetingDateOf(md || {}, json),
    };
  }

  // Returns captures. Optional sourceUrl supplies drive/item ids when the JSON
  // itself has none (collection list, cdnmedia, item GET).
  function scanForTranscripts(json, sourceUrl) {
    const found = [];
    if (!json || typeof json !== 'object') return found;
    const fromSrc = sourceUrl ? parseSharePointApiUrl(sourceUrl) : null;
    const rootItemName = str(json.name);
    const rootCreated = str(json.createdDateTime);

    (function walk(node, path, parentItem) {
      if (!node || typeof node !== 'object' || path.length > MAX_DEPTH) return;
      if (Array.isArray(node)) {
        node.forEach((child, i) => walk(child, path.concat(String(i)), parentItem));
        return;
      }
      const item = str(node.name) && str(node.createdDateTime) ? node : parentItem;
      if (looksLikeTranscript(node, path, fromSrc)) {
        const cap = captureFrom(node, path, item, rootItemName, rootCreated, json, fromSrc);
        if (cap && (cap.url || cap.contentUrl || (cap.driveId && cap.itemId))) found.push(cap);
      }
      for (const key of Object.keys(node)) {
        const v = node[key];
        if (v && typeof v === 'object') walk(v, path.concat(key), item);
      }
    })(json, [], undefined);

    if (!found.length && fromSrc && looksLikeRecordingItem(json, sourceUrl)) {
      found.push({
        name: rootItemName || 'transcript',
        url: buildTranscriptUrl(fromSrc, 'list') || sourceUrl.split('?')[0],
        driveId: fromSrc.driveId,
        itemId: fromSrc.itemId,
        transcriptId: fromSrc.transcriptId,
        host: fromSrc.host,
        origin: fromSrc.origin,
        sitePath: fromSrc.sitePath,
        itemName: rootItemName,
        createdDateTime: rootCreated,
        meetingDate: meetingDateOf(json, json),
      });
    }

    const seen = new Set();
    return found.filter((t) => {
      const k = (t.contentUrl || t.url || `${t.driveId}/${t.itemId}/${t.transcriptId || ''}`).split('?')[0];
      if (seen.has(k)) return false;
      seen.add(k);
      return true;
    });
  }

  function looksLikeRecordingItem(json, url) {
    if (url && /transcript/i.test(url)) return true;
    if (!json || typeof json !== 'object') return false;
    if (json.media) return true;
    if (json.file && /video/i.test(String(json.file.mimeType || ''))) return true;
    if (/meeting recording/i.test(String(json.name || ''))) return true;
    return false;
  }

  function captureFromRequestUrl(url, fileInfo) {
    const p = parseSharePointApiUrl(url);
    if (!p) return null;
    const name = (fileInfo && (str(fileInfo.displayName) || str(fileInfo.name))) || 'transcript';
    const contentUrl = p.transcriptId ? buildTranscriptUrl(p, 'content') : undefined;
    return {
      name,
      url: contentUrl || buildTranscriptUrl(p, 'list') || url.split('?')[0],
      contentUrl,
      driveId: p.driveId,
      itemId: p.itemId,
      transcriptId: p.transcriptId,
      host: p.host,
      origin: p.origin,
      sitePath: p.sitePath,
      itemName: name,
    };
  }

  // Prefer site-path form first (matches Recap), then tenant-root `_api/v2.1`.
  function altParts(p) {
    if (!p) return [];
    const out = [p];
    if (p.sitePath) out.push(Object.assign({}, p, { sitePath: '' }));
    return out;
  }

  function downloadCandidates(t) {
    const out = [];
    const seen = new Set();
    const add = (u) => {
      if (!u || seen.has(u)) return;
      seen.add(u);
      out.push(u);
    };
    const addSharePoint = (p) => {
      if (!p || !p.transcriptId) return;
      add(buildTranscriptUrl(p, 'content', 'json'));
      add(withEditsOff(buildTranscriptUrl(p, 'content', 'json')));
      add(buildTranscriptUrl(p, 'streamContent', 'json'));
      add(withEditsOff(buildTranscriptUrl(p, 'streamContent', 'json')));
      add(buildTranscriptUrl(p, 'content'));
      add(buildTranscriptUrl(p, 'streamContent'));
    };
    const parsed = partsFromCapture(t);
    altParts(parsed).forEach(addSharePoint);
    // Collab / mcps location (Teams page session). Do not require recording Download.
    if (t && t.location && !isCdnOrTempauth(t.location)) {
      add(withFormat(t.location, 'json'));
      add(withEditsOff(withFormat(t.location, 'json')));
      add(t.location);
    }
    if (t && t.contentUrl && !isCdnOrTempauth(t.contentUrl)) {
      add(withFormat(t.contentUrl, 'json'));
      add(t.contentUrl);
    }
    if (t && t.url && !isCdnOrTempauth(t.url)) {
      add(withFormat(t.url, 'json'));
      add(t.url);
    }
    // Last resort: the original tempauth link, in case cookie-auth is blocked.
    if (t && t.url && isCdnOrTempauth(t.url) && urlKind(t.url) === 'tempauth') {
      add(withFormat(t.url, 'json'));
      add(t.url);
    }
    return out;
  }

  function discoverUrls(t) {
    const parsed = partsFromCapture(t);
    const out = [];
    const seen = new Set();
    const add = (u) => {
      if (!u || seen.has(u)) return;
      seen.add(u);
      out.push(u);
    };
    altParts(parsed).forEach((p) => {
      add(buildTranscriptUrl(p, 'expand'));
      add(buildTranscriptUrl(p, 'list'));
      add(buildTranscriptUrl(p, 'listCurrent'));
      if (p.transcriptId) add(buildTranscriptUrl(p, 'metadata'));
    });
    if (t && t.location && (isCollabUrl(t.location) || parseSharePointApiUrl(t.location))) add(t.location);
    return out;
  }

  function debugCapture(t) {
    return {
      name: t && t.name,
      language: t && t.language,
      transcriptType: t && t.transcriptType,
      capturedAt: t && t.capturedAt,
      createdDateTime: t && t.createdDateTime,
      meetingDate: t && t.meetingDate,
      hasUrl: !!(t && t.url),
      hasContentUrl: !!(t && t.contentUrl),
      hasDriveItem: !!(t && t.driveId && t.itemId),
      hasTranscriptId: !!(t && t.transcriptId),
      hasLocation: !!(t && t.location),
      urlKind: urlKind(t && (t.contentUrl || t.url || t.location)),
    };
  }

  // Cheap pre-check so we only JSON.parse responses that can contain a link.
  function mightContainTranscript(text) {
    if (typeof text !== 'string') return false;
    return text.indexOf('temporaryDownloadUrl') !== -1
      || text.indexOf('transcriptId') !== -1
      || text.indexOf('speakerDisplayName') !== -1
      || text.indexOf('"transcriptType"') !== -1
      || text.indexOf('MeetingTranscript') !== -1
      || text.indexOf('/media/transcripts') !== -1
      || text.indexOf('cdnmedia/transcripts') !== -1
      || text.indexOf('stream.office.com/schemas/transcript') !== -1;
  }

  function isTranscriptishUrl(url) {
    if (typeof url !== 'string') return false;
    return /\/(?:cdnmedia|media)\/transcripts/i.test(url)
      || /[?&]\$?expand=.*transcript/i.test(url);
  }

  function isSharePointItemUrl(url) {
    return typeof url === 'string' && /\/_api(?:_cached)?\/v[0-9.]+\/drives\/[^/]+\/items\/[^/]+/i.test(url);
  }

  const api = {
    scanForTranscripts,
    mightContainTranscript,
    parseSharePointApiUrl,
    parseRecapHref,
    buildTranscriptUrl,
    downloadCandidates,
    discoverUrls,
    partsFromCapture,
    captureFromRequestUrl,
    debugCapture,
    urlKind,
    ticksToIso,
    withFormat,
    looksLikeTranscriptDoc,
    isTranscriptishUrl,
    isSharePointItemUrl,
    looksLikeRecordingItem,
    isCollabUrl,
  };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else Object.defineProperty(root, '__teamsTranscriptScan', { value: api, configurable: false });
})(typeof globalThis !== 'undefined' ? globalThis : this);
