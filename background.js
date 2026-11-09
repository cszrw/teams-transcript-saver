// Service worker: keeps per-tab captures (in memory-only session storage),
// downloads and saves transcripts. Prefers cookie-auth SharePoint /content
// over short-lived tempauth links (those now 401/403 on current Teams recap).
importScripts('lib/convert.js', 'lib/naming.js', 'lib/scan.js');

const MAX_PER_TAB = 10;
const ALLOWED_HOST = /(^|\.)(sharepoint\.com(?:\.mcas\.ms)?|svc\.ms|teams\.microsoft\.com(?:\.mcas\.ms)?|teams\.cloud\.microsoft)$/i;
const key = (tabId) => `tab:${tabId}`;
const Scan = self.__teamsTranscriptScan;
const lastTopPath = new Map();
const discovering = new Set();

async function getCaptures(tabId) {
  const k = key(tabId);
  return (await chrome.storage.session.get(k))[k] || [];
}
async function setCaptures(tabId, list) {
  await chrome.storage.session.set({ [key(tabId)]: list });
  await updateUi(tabId, list.length);
}

async function updateUi(tabId, count) {
  await chrome.action.setBadgeText({ tabId, text: count ? String(count) : '' }).catch(() => {});
  await chrome.action.setBadgeBackgroundColor({ tabId, color: '#5b5fc7' }).catch(() => {});
  chrome.tabs.sendMessage(tabId, { type: 'state', count }, { frameId: 0 }).catch(() => {});
}

function isAllowedUrl(url) {
  try {
    const u = new URL(url);
    return u.protocol === 'https:' && ALLOWED_HOST.test(u.hostname);
  } catch (e) { return false; }
}

const transcriptKey = (t) => {
  if (t.contentUrl) return t.contentUrl.split('?')[0];
  if (t.driveId && t.itemId) return `${t.driveId}/${t.itemId}/${t.transcriptId || ''}`;
  if (t.url) return t.url.split('?')[0];
  return t.id;
};

async function addCaptures(tabId, frameId, transcripts) {
  let list = await getCaptures(tabId);
  for (const t of transcripts) {
    if (!t) continue;
    const url = t.contentUrl || t.url;
    if (url && !isAllowedUrl(url)) continue;
    if (!url && !(t.driveId && t.itemId && (t.host || t.origin))) continue;
    const k = transcriptKey(t);
    list = list.filter((x) => transcriptKey(x) !== k);
    list.push({ ...t, frameId, id: k });
  }
  await setCaptures(tabId, list.slice(-MAX_PER_TAB));
}

class ExpiredError extends Error {}

async function fetchText(tabId, frameId, url) {
  if (!isAllowedUrl(url)) return { ok: false, status: 0, error: 'blocked host' };
  let r = await chrome.tabs.sendMessage(tabId, { type: 'page-fetch', url }, { frameId }).catch(() => null);
  if (!r || (!r.ok && !r.status)) {
    try {
      const res = await fetch(url, { credentials: 'include', cache: 'no-store' });
      r = { ok: res.ok, status: res.status, text: res.ok ? await res.text() : '' };
    } catch (e) {
      r = { ok: false, status: 0, error: String(e.message || e) };
    }
  }
  return r;
}

function usableBody(text, wantJson) {
  if (!text || !String(text).trim()) return false;
  const ext = TranscriptConvert.detectExtension(text);
  if (wantJson && ext !== 'json') return false;
  if (ext === 'txt' && !TranscriptConvert.toReadable(text)) return false;
  return ext === 'json' || ext === 'vtt' || !!TranscriptConvert.toReadable(text);
}

async function refreshTranscript(tabId, t) {
  const urls = Scan.discoverUrls(t);
  for (const url of urls) {
    const r = await fetchText(tabId, t.frameId || 0, url);
    if (!r || !r.ok || !r.text) continue;
    let json;
    try { json = JSON.parse(r.text); } catch (e) { continue; }
    const found = Scan.scanForTranscripts(json, url);
    if (found.length) {
      await addCaptures(tabId, t.frameId, found);
      return found[found.length - 1];
    }
  }
  const parsed = Scan.partsFromCapture(t);
  if (parsed && parsed.transcriptId) {
    const contentUrl = Scan.buildTranscriptUrl(parsed, 'content');
    if (contentUrl) return { ...t, url: contentUrl, contentUrl };
  }
  return null;
}

async function download(tabId, t) {
  let current = t;
  if (!(current.transcriptId || (Scan.partsFromCapture(current) || {}).transcriptId)) {
    const listed = await refreshTranscript(tabId, current);
    if (listed) current = { ...current, ...listed };
  }
  const tried = new Set();
  async function attempt(src) {
    let last;
    for (const url of Scan.downloadCandidates(src)) {
      if (tried.has(url)) continue;
      tried.add(url);
      last = await fetchText(tabId, src.frameId || 0, url);
      const wantJson = /[?&]format=json(?:&|$)/.test(url);
      if (last && last.ok && usableBody(last.text, wantJson)) return { text: last.text, last };
      if (last && last.ok && last.text && wantJson) continue;
    }
    return { text: null, last };
  }
  let { text, last } = await attempt(current);
  const expired = last && (last.status === 401 || last.status === 403);
  if (!text && expired) {
    const fresh = await refreshTranscript(tabId, current);
    if (fresh) {
      current = { ...current, ...fresh };
      ({ text, last } = await attempt(current));
    }
  }
  if (text) return text;
  if (last && (last.status === 401 || last.status === 403)) throw new ExpiredError();
  throw new Error(`Download failed (${last && (last.status || last.error) || 'no response'}).`);
}

function dataUrl(text, mime) {
  const bytes = new TextEncoder().encode(text);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return `data:${mime};charset=utf-8;base64,${btoa(bin)}`;
}

function saveFile(name, text, mime) {
  return chrome.downloads.download({
    url: dataUrl(text, mime),
    filename: `teams-transcripts/${name}`,
    conflictAction: 'uniquify',
    saveAs: false,
  });
}

const MIME = { json: 'application/json', vtt: 'text/vtt', txt: 'text/plain' };

function notesFor(t, extra) {
  const notes = extra ? extra.slice() : [];
  if (!t.meetingDate) notes.push('transcript capture had no meetingStartDateTime');
  if (!t.transcriptId && !(t.driveId && t.itemId)) notes.push('no SharePoint drive/item ids on the capture; reopen Recap → Transcript');
  return notes;
}

async function saveDebug(base, payload) {
  await saveFile(`${base}_debug.json`, JSON.stringify(payload, null, 2) + '\n', MIME.json);
}

async function saveTranscript(tabId, t) {
  const base = TranscriptNaming.baseName(t);
  const debug = {
    version: chrome.runtime.getManifest().version,
    savedAt: new Date().toISOString(),
    notes: notesFor(t),
    capture: Scan.debugCapture(t),
    result: { ok: false },
    files: [`${base}_debug.json`],
  };
  try {
    const raw = await download(tabId, t);
    const ext = TranscriptConvert.detectExtension(raw);
    await saveFile(`${base}.${ext}`, raw, MIME[ext] || MIME.txt);
    debug.files.push(`${base}.${ext}`);
    const readable = ext === 'txt' ? null : TranscriptConvert.toReadable(raw);
    if (readable) {
      const header = `${TranscriptNaming.meetingTitle(t)} (${base.slice(0, 10)})\n\n`;
      await saveFile(`${base}.txt`, header + readable.lines.join('\n') + '\n', MIME.txt);
      debug.files.push(`${base}.txt`);
    }
    debug.result = {
      ok: true,
      lines: readable ? readable.lines.length : 0,
      hasSpeakers: !!(readable && readable.hasSpeakers),
    };
    if (debug.result.lines && !debug.result.hasSpeakers) debug.notes.push('no speaker names in this transcript');
    await saveDebug(base, debug);
    return {
      files: debug.files.filter((f) => !f.endsWith('_debug.json')).concat([`${base}_debug.json`]),
      lines: debug.result.lines,
      hasSpeakers: debug.result.hasSpeakers,
      base,
    };
  } catch (e) {
    debug.result = { ok: false, error: e instanceof ExpiredError ? 'expired' : 'failed', message: e.message };
    if (e instanceof ExpiredError) {
      debug.notes.push('save failed: download link expired (401/403); retried cookie-auth /content');
    }
    await saveDebug(base, debug);
    throw e;
  }
}

async function handleSave(tabId, id) {
  const list = await getCaptures(tabId);
  const t = id ? list.find((x) => x.id === id) : list[list.length - 1];
  if (!t) return { ok: false, message: 'No transcript captured on this tab yet. Open the meeting\'s Recap → Transcript.' };
  try {
    const r = await saveTranscript(tabId, t);
    const note = r.lines && !r.hasSpeakers ? ' (no speaker names in this transcript)' : '';
    return { ok: true, message: `Saved to Downloads/teams-transcripts: ${r.files.join(', ')}${note}` };
  } catch (e) {
    if (e instanceof ExpiredError) {
      return { ok: false, message: 'The download link has expired. Close and reopen the transcript in Teams, then try again.' };
    }
    return { ok: false, message: `${e.message} Reopen the transcript in Teams (or reload the page) and try again.` };
  }
}

async function discoverFromItem(tabId, frameId, item) {
  if (!item || !item.driveId || !item.itemId) return;
  const k = `${tabId}:${item.driveId}:${item.itemId}`;
  if (discovering.has(k)) return;
  discovering.add(k);
  try {
    const fake = {
      driveId: item.driveId,
      itemId: item.itemId,
      transcriptId: item.transcriptId,
      host: item.host,
      origin: item.origin,
      sitePath: item.sitePath,
      name: item.name,
    };
    const found = [];
    for (const url of Scan.discoverUrls(fake)) {
      if (!isAllowedUrl(url)) continue;
      const r = await fetchText(tabId, frameId || 0, url);
      if (!r || !r.ok || !r.text) continue;
      let json;
      try { json = JSON.parse(r.text); } catch (e) { continue; }
      found.push.apply(found, Scan.scanForTranscripts(json, url));
      if (found.length) break;
    }
    if (found.length) await addCaptures(tabId, frameId, found);
    else await addCaptures(tabId, frameId, [fake]);
  } finally {
    discovering.delete(k);
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const senderTab = sender.tab && sender.tab.id;
  if (msg.type === 'captured' && senderTab != null) {
    addCaptures(senderTab, sender.frameId, msg.transcripts || []);
    return false;
  }
  if (msg.type === 'item-context' && senderTab != null) {
    discoverFromItem(senderTab, sender.frameId, msg.item).catch(() => {});
    return false;
  }
  if (msg.type === 'hello' && senderTab != null) {
    getCaptures(senderTab).then((l) => sendResponse({ count: l.length }));
    return true;
  }
  if (msg.type === 'save-latest' && senderTab != null) {
    handleSave(senderTab).then((r) => {
      chrome.tabs.sendMessage(senderTab, { type: 'status', text: r.message, error: !r.ok }, { frameId: 0 }).catch(() => {});
    });
    return false;
  }
  if (msg.type === 'list') {
    getCaptures(msg.tabId).then((l) => sendResponse(l.map(({ id, name, language, transcriptType, capturedAt, createdDateTime }) =>
      ({ id, name, language, transcriptType, capturedAt, createdDateTime }))));
    return true;
  }
  if (msg.type === 'save') {
    handleSave(msg.tabId, msg.id).then(sendResponse);
    return true;
  }
  return false;
});

chrome.tabs.onRemoved.addListener((tabId) => {
  chrome.storage.session.remove(key(tabId));
  lastTopPath.delete(tabId);
});
chrome.tabs.onUpdated.addListener((tabId, info) => {
  // Only drop captures when the top-level path actually changes. Teams recap
  // is a SPA: status=loading fires as iframes load and used to wipe a still-valid capture.
  if (!info.url) return;
  let next;
  try {
    const u = new URL(info.url);
    next = u.origin + u.pathname;
  } catch (e) { return; }
  const prev = lastTopPath.get(tabId);
  lastTopPath.set(tabId, next);
  if (prev && prev !== next) {
    chrome.storage.session.remove(key(tabId)).then(() => updateUi(tabId, 0));
  }
});
