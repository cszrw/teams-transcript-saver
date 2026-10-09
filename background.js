// Service worker: keeps per-tab captures (in memory-only session storage),
// downloads and saves transcripts.
importScripts('lib/convert.js', 'lib/naming.js');

const MAX_PER_TAB = 10;
const ALLOWED_HOST = /(^|\.)(sharepoint\.com|svc\.ms|teams\.microsoft\.com|teams\.cloud\.microsoft)$/i;
const key = (tabId) => `tab:${tabId}`;

// chrome.storage.session lives in memory only and is wiped when Chrome closes.
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
  try { const u = new URL(url); return u.protocol === 'https:' && ALLOWED_HOST.test(u.hostname); } catch (e) { return false; }
}

// Same transcript, refreshed token -> same key (URL without its query string).
const transcriptKey = (url) => url.split('?')[0];

async function addCaptures(tabId, frameId, transcripts) {
  let list = await getCaptures(tabId);
  for (const t of transcripts) {
    if (!t || !isAllowedUrl(t.url)) continue;
    const k = transcriptKey(t.url);
    list = list.filter((x) => transcriptKey(x.url) !== k);
    list.push({ ...t, frameId, id: k });
  }
  await setCaptures(tabId, list.slice(-MAX_PER_TAB));
}

function withFormatJson(url) {
  const u = new URL(url);
  u.searchParams.set('format', 'json');
  return u.toString();
}

class ExpiredError extends Error {}

// Fetch from the page first (same session as Teams), then from the service worker.
async function fetchText(tabId, frameId, url) {
  let r = await chrome.tabs.sendMessage(tabId, { type: 'page-fetch', url }, { frameId }).catch(() => null);
  if (!r || (!r.ok && !r.status)) {
    try {
      const res = await fetch(url, { credentials: 'include', cache: 'no-store' });
      r = { ok: res.ok, status: res.status, text: res.ok ? await res.text() : '' };
    } catch (e) {
      r = { ok: false, status: 0, error: String(e.message || e) };
    }
  }
  if (r.status === 401 || r.status === 403) throw new ExpiredError();
  return r;
}

// Try JSON first (has speaker names), then the plain stream (usually WEBVTT).
async function download(tabId, t) {
  const tries = [withFormatJson(t.url), t.url];
  let last;
  for (const url of tries) {
    last = await fetchText(tabId, t.frameId || 0, url);
    if (last.ok && last.text && last.text.trim()) {
      if (url !== t.url && TranscriptConvert.detectExtension(last.text) !== 'json') continue;
      return last.text;
    }
  }
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

async function saveTranscript(tabId, t) {
  const raw = await download(tabId, t);
  const ext = TranscriptConvert.detectExtension(raw);
  const base = TranscriptNaming.baseName(t);
  await saveFile(`${base}.${ext}`, raw, MIME[ext]);
  const files = [`${base}.${ext}`];
  const readable = ext === 'txt' ? null : TranscriptConvert.toReadable(raw);
  if (readable) {
    const header = `${TranscriptNaming.meetingTitle(t)} (${base.slice(0, 10)})\n\n`;
    await saveFile(`${base}.txt`, header + readable.lines.join('\n') + '\n', MIME.txt);
    files.push(`${base}.txt`);
  }
  return { files, lines: readable ? readable.lines.length : 0, hasSpeakers: !!(readable && readable.hasSpeakers) };
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

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const senderTab = sender.tab && sender.tab.id;
  if (msg.type === 'captured' && senderTab != null) {
    addCaptures(senderTab, sender.frameId, msg.transcripts || []);
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
  // From the popup:
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

chrome.tabs.onRemoved.addListener((tabId) => chrome.storage.session.remove(key(tabId)));
chrome.tabs.onUpdated.addListener((tabId, info) => {
  // A full page reload starts fresh; links from the old page would be stale anyway.
  if (info.status === 'loading') chrome.storage.session.remove(key(tabId)).then(() => updateUi(tabId, 0));
});
