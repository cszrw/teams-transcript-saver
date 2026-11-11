// Service worker: keeps per-tab captures (in memory-only session storage),
// downloads and saves transcripts.
importScripts('lib/convert.js', 'lib/naming.js', 'lib/output.js');

const MAX_PER_TAB = 10;
const ALLOWED_HOST = /(^|\.)(sharepoint\.com|svc\.ms|teams\.microsoft\.com|teams\.cloud\.microsoft)$/i;
const key = (tabId) => `tab:${tabId}`;

// chrome.storage.session lives in memory only and is wiped when Chrome closes.
async function getCaptures(tabId) {
  const k = key(tabId);
  return (await chrome.storage.session.get(k))[k] || [];
}
async function setCaptures(tabId, list) {
  try { await chrome.storage.session.set({ [key(tabId)]: list }); }
  catch (e) {
    // Quota trouble must never cost the button: drop the optional meeting data and retry.
    console.log('[TTS] storing captures failed, clearing meeting data', e);
    await chrome.storage.session.remove(metaKey(tabId)).catch(() => {});
    await chrome.storage.session.set({ [key(tabId)]: list }).catch(() => {});
  }
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

// Every awaited lookup gets a deadline so a save can never hang on 'Saving…'.
function withTimeout(promise, ms, label) {
  let timer;
  return Promise.race([
    Promise.resolve(promise).finally(() => clearTimeout(timer)),
    new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(`${label} timed out after ${ms / 1000}s.`)), ms); }),
  ]);
}
async function getMetaSafe(tabId) {
  try { return await withTimeout(getMeta(tabId), 5000, 'meeting info'); }
  catch (e) { console.log('[TTS]', e.message); return { meetings: [], chats: {} }; }
}

// Fetch from the page first (same session as Teams), then from the service worker.
async function fetchText(tabId, frameId, url) {
  let r = await withTimeout(chrome.tabs.sendMessage(tabId, { type: 'page-fetch', url }, { frameId }), 15000, 'page download').catch((e) => { console.log('[TTS]', e.message); return null; });
  if (!r || (!r.ok && !r.status)) {
    const ctl = new AbortController();
    const timer = setTimeout(() => ctl.abort(), 15000);
    try {
      const res = await fetch(url, { credentials: 'include', cache: 'no-store', signal: ctl.signal });
      r = { ok: res.ok, status: res.status, text: res.ok ? await res.text() : '' };
    } catch (e) {
      r = { ok: false, status: 0, error: ctl.signal.aborted ? 'timed out' : String(e.message || e) };
    } finally { clearTimeout(timer); }
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
  return withTimeout(chrome.downloads.download({
    url: dataUrl(text, mime),
    filename: `teams-transcripts/${name}`,
    conflictAction: 'overwrite',
    saveAs: false,
  }), 10000, `saving ${name}`);
}

const MIME = { json: 'application/json', vtt: 'text/vtt', txt: 'text/plain' };

// ---- meeting metadata + chat seen on this tab (memory-only, like captures) ----
const metaKey = (tabId) => `meta:${tabId}`;
async function getMeta(tabId) {
  const k = metaKey(tabId);
  return (await chrome.storage.session.get(k))[k] || { meetings: [], chats: {} };
}
let metaQueue = Promise.resolve();
function addMeta(tabId, meetings, chat) {
  metaQueue = metaQueue.then(async () => {
    const m = await getMeta(tabId);
    for (const x of meetings || []) {
      m.meetings = m.meetings.filter((y) => !(y.start === x.start && y.subject === x.subject));
      m.meetings.push({ ...x, seenAt: Date.now() });
    }
    m.meetings = m.meetings.slice(-50);
    if (chat && chat.threadId) {
      const c = m.chats[chat.threadId] || { messages: [], callStart: undefined };
      const byId = new Map(c.messages.map((x) => [x.id, x]));
      for (const x of chat.messages) byId.set(x.id, x);
      c.messages = [...byId.values()].sort((a, b) => (a.time < b.time ? -1 : 1)).slice(-500);
      if (chat.callStart && (!c.callStart || chat.callStart < c.callStart)) c.callStart = chat.callStart;
      c.seenAt = Date.now();
      m.chats[chat.threadId] = c;
      const ids = Object.keys(m.chats).sort((a, b) => m.chats[b].seenAt - m.chats[a].seenAt);
      for (const id of ids.slice(5)) delete m.chats[id]; // keep storage small
    }
    await chrome.storage.session.set({ [metaKey(tabId)]: m });
  }).catch((e) => console.log('[TTS] meta store failed', e));
}

function safe(fn, fallback) { try { return fn(); } catch (e) { console.log('[TTS]', e); return fallback; } }

// ---- debug: every JSON response the page script saw on this tab ----
// Kept in service-worker memory only (never chrome.storage: 400 summaries filled the 10 MB
// session quota, so storing the transcript captures failed and the button never appeared).
const observations = new Map(); // tabId -> [summary]
const MAX_OBS = 200;
function addObservation(tabId, frameId, entry) {
  try {
    if (!entry || typeof entry !== 'object') return;
    const list = observations.get(tabId) || [];
    if (list.length >= MAX_OBS) list.shift();
    list.push(Object.assign({ frameId }, entry));
    observations.set(tabId, list);
  } catch (e) { console.log('[TTS] observe failed', e); }
}

async function saveDebug(tabId, base, t, info, notes) {
  try {
    const observed = observations.get(tabId) || [];
    const meta = await getMetaSafe(tabId);
    const dbg = {
      version: chrome.runtime.getManifest().version, savedAt: new Date().toISOString(),
      transcript: { name: t.name, itemName: t.itemName, createdDateTime: t.createdDateTime, meetingDate: t.meetingDate, language: t.language, capturedAt: t.capturedAt },
      resolved: { start: info.start || null, startSource: info.startSource || null, subject: info.subject || null, threadId: info.threadId || null, chatMessages: info.chat ? info.chat.messages.length : 0 },
      notes,
      meetingsSeen: (meta.meetings || []).map(({ start, end, subject, threadId }) => ({ start, end, subject, threadId })),
      chatsSeen: Object.entries(meta.chats || {}).map(([id, c]) => ({ threadId: id, messages: c.messages.length, callStart: c.callStart })),
      observedCount: observed.length, observed,
    };
    await saveFile(`${base}_debug.json`, JSON.stringify(dbg, null, 1), MIME.json);
    return `${base}_debug.json`;
  } catch (e) { console.log('[TTS] debug dump failed', e); return null; }
}

async function saveTranscript(tabId, t) {
  const files = [];
  const notes = [];
  const title = safe(() => TranscriptNaming.meetingTitle(t), 'transcript');
  let info = {};
  let base = safe(() => TranscriptNaming.baseName(t), 'transcript');
  let readable = null;
  try {
    const meta = await getMetaSafe(tabId);
    try { info = TranscriptOutput.resolve(t, meta, title) || {}; } catch (e) { notes.push('meeting info failed: ' + e.message); info = {}; }
    base = safe(() => TranscriptNaming.baseName({ ...t, meetingDate: info.start || t.meetingDate }), base);
    if (!t.meetingDate) notes.push('transcript capture had no meetingStartDateTime');
    if (!(meta.meetings || []).length) notes.push('no meeting-details candidates (start + subject/thread id) seen on this tab');
    else if (!info.start) notes.push(`meeting candidates seen, none matching "${title}" and not all from one meeting thread`);
    if (info.threadId && !info.chat) notes.push('meeting thread known but no chat messages were loaded by the page (open the meeting Chat tab before saving)');
    if (!info.threadId) notes.push('no meeting chat thread (19:meeting_…@thread.v2) linked');

    const raw = await withTimeout(download(tabId, t), 40000, 'transcript download');
    const ext = TranscriptConvert.detectExtension(raw);
    readable = ext === 'txt' ? null : safe(() => TranscriptConvert.toReadable(raw), null);
    await saveFile(`${base}.${ext}`, ext === 'json' ? safe(() => TranscriptOutput.augmentJson(raw, info), raw) : raw, MIME[ext]);
    files.push(`${base}.${ext}`);
    if (readable) {
      const header = safe(() => TranscriptOutput.header(title, base.slice(0, 10), info), `${title} (${base.slice(0, 10)})\n\n`);
      await saveFile(`${base}.txt`, header + readable.lines.join('\n') + '\n' + safe(() => TranscriptOutput.chatSection(info), ''), MIME.txt);
      files.push(`${base}.txt`);
    }
  } catch (e) {
    notes.push('save failed: ' + (e instanceof ExpiredError ? 'download link expired (401/403)' : e.message));
    e.files = files;
    e.debugFile = await saveDebug(tabId, base, t, info, notes);
    throw e;
  }
  const dbgFile = await saveDebug(tabId, base, t, info, notes);
  if (dbgFile) files.push(dbgFile);
  if (!info.start) files.push('(meeting date not found, download date used)');
  if (info.chat) files.push(`(${info.chat.messages.length} chat messages)`);
  return { files, lines: readable ? readable.lines.length : 0, hasSpeakers: !!(readable && readable.hasSpeakers) };
}

async function handleSave(tabId, id) {
  let list = [];
  try { list = await withTimeout(getCaptures(tabId), 5000, 'reading captures'); } catch (e) { return { ok: false, message: e.message }; }
  const t = id ? list.find((x) => x.id === id) : list[list.length - 1];
  if (!t) return { ok: false, message: 'No transcript captured on this tab yet. Open the meeting\'s Recap → Transcript.' };
  try {
    const r = await saveTranscript(tabId, t);
    const note = r.lines && !r.hasSpeakers ? ' (no speaker names in this transcript)' : '';
    return { ok: true, message: `Saved to Downloads/teams-transcripts: ${r.files.join(', ')}${note}` };
  } catch (e) {
    const dbg = e && e.debugFile ? ` Debug info: ${e.debugFile}.` : '';
    if (e instanceof ExpiredError) {
      return { ok: false, message: `The download link has expired. Close and reopen the transcript in Teams, then try again.${dbg}` };
    }
    return { ok: false, message: `${e.message}${dbg} Reopen the transcript in Teams (or reload the page) and try again.` };
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const senderTab = sender.tab && sender.tab.id;
  if (msg.type === 'captured' && senderTab != null) {
    addCaptures(senderTab, sender.frameId, msg.transcripts || []);
    return false;
  }
  if (msg.type === 'observe' && senderTab != null) {
    addObservation(senderTab, sender.frameId, msg.entry);
    return false;
  }
  if (msg.type === 'meta' && senderTab != null) {
    try { addMeta(senderTab, msg.meetings, msg.chat); } catch (e) { console.log('[TTS] meta', e); }
    return false;
  }
  if (msg.type === 'hello' && senderTab != null) {
    if (sender.frameId === 0 && msg.fresh) {
      observations.delete(senderTab);
      chrome.storage.session.remove([key(senderTab), metaKey(senderTab)]).catch(() => {})
        .then(() => getCaptures(senderTab)).then((l) => sendResponse({ count: l.length }));
      return true;
    }
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

chrome.tabs.onRemoved.addListener((tabId) => { observations.delete(tabId); chrome.storage.session.remove([key(tabId), metaKey(tabId)]); });
// A new top-level document (full reload) starts fresh; that is signalled by the top frame's 'hello'.
// (tabs.onUpdated status 'loading' also fires for Teams' in-page navigations and wiped the button.)
