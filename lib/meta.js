// Finds meeting metadata (start time, subject, chat thread) and chat messages
// in JSON API responses the Teams page already loads. Shared by inject.js and tests.
(function (root) {
  'use strict';
  const MAX_DEPTH = 10;
  const str = (v) => (typeof v === 'string' && v ? v : undefined);
  const isDate = (v) => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v) && !isNaN(new Date(v));
  // Graph/Outlook style {dateTime, timeZone}: treat as UTC when timeZone is UTC or missing.
  function dt(v) {
    if (isDate(v)) return new Date(v).toISOString();
    if (v && typeof v === 'object' && isDate(v.dateTime)) {
      const s = v.dateTime;
      if (/Z$|[+-]\d\d:?\d\d$/.test(s)) return new Date(s).toISOString();
      const utc = !v.timeZone || /^(UTC|GMT|Etc\/UTC|Coordinated Universal Time)$/i.test(v.timeZone);
      return new Date(utc ? s + 'Z' : s).toISOString(); // non-UTC zone: assume browser-local (Berlin)
    }
    return undefined;
  }
  const THREAD = /19:meeting_[A-Za-z0-9_-]+@thread\.v2/;

  // Returns [{ start, end, subject, threadId, source }]
  // .NET ticks (100 ns since 0001-01-01), e.g. readcollabobject metadata.startTime "639279937919530198".
  function ticks(v) {
    if (!(typeof v === 'string' && /^6\d{17}$/.test(v)) && !(typeof v === 'number' && v > 6e17)) return undefined;
    const ms = (Number(v) - 621355968000000000) / 1e4;
    return ms > 1.5e12 && ms < 2.5e12 ? new Date(ms).toISOString() : undefined;
  }
  const anyDt = (...vals) => { for (const v of vals) { const d = dt(v) || ticks(v); if (d) return d; } return undefined; };
  const threadOf = (...vals) => {
    for (const x of vals) { const m = typeof x === 'string' ? (() => { try { return decodeURIComponent(x); } catch (e) { return x; } })().match(THREAD) : null; if (m) return m[0]; }
    return undefined;
  };

  // Returns [{ start, end, subject, threadId, source }]
  function scanMeetings(json, diag) {
    const out = [];
    (function walk(n, depth) {
      if (!n || typeof n !== 'object' || depth > MAX_DEPTH) return;
      if (Array.isArray(n)) { n.forEach((c) => walk(c, depth + 1)); return; }
      const md = n.metadata && typeof n.metadata === 'object' ? n.metadata : {};
      const media = n.media && typeof n.media === 'object' ? n.media : {};
      // Recording driveItem: media.recordingStartDateTime + source.threadId (organiser's OneDrive).
      const recStart = anyDt(media.recordingStartDateTime, media.meetingStartDateTime);
      const start = recStart || anyDt(n.meetingStartDateTime, n.MeetingStartTime, n.RecordingStartTime, n.startDateTime,
        n.scheduledStartTime, n.callStartTime, n.startTime, n.start, md.startTime, md.meetingStartTime);
      if (start) {
        const subject = str(n.subject) || str(n.Subject) || str(n.meetingTitle) || str(n.title) || str(n.Title) || str(n.topic) ||
          (recStart ? undefined : str(n.displayName)) || str(md.title) || str(md.subject);
        const tid = threadOf(n.threadId, n.chatId, md.threadId, n.source && n.source.threadId,
          n.MeetingRecording && n.MeetingRecording.TeamsThreadId, n.TeamsChatUrl, n.joinWebUrl, n.joinUrl,
          n.onlineMeeting && n.onlineMeeting.joinUrl, n.chatInfo && n.chatInfo.threadId);
        if (!(subject || tid) && diag) diag.push(`meeting candidate rejected (start ${start}, no subject/thread id); keys: ${Object.keys(n).slice(0, 12).join(',')}`);
        if (subject || tid) {
          out.push({ start,
            end: anyDt(media.recordingEndDateTime, n.meetingEndDateTime, n.MeetingEndTime, n.RecordingEndTime, n.endDateTime,
              n.scheduledEndTime, n.endTime, n.end, md.endTime),
            subject, threadId: tid, source: recStart ? 'recording' : md.startTime ? 'collab' : n.MeetingStartTime ? 'search' : 'meeting' });
          if (recStart || md.startTime) return; // don't also emit nested pieces of the same object
        }
      }
      for (const k of Object.keys(n)) if (n[k] && typeof n[k] === 'object') walk(n[k], depth + 1);
    })(json, 0);
    return out;
  }

  function htmlToText(h) {
    return String(h || '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div)>/gi, '\n').replace(/<[^>]+>/g, '')
      .replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
      .replace(/\n{3,}/g, '\n\n').trim();
  }

  // Teams chat service (.../conversations/<thread>/messages) and Graph chat messages.
  // Returns { threadId, messages: [{ id, time, sender, text }], callStart }
  function scanChat(json, diag) {
    const list = json && (Array.isArray(json.messages) ? json.messages : Array.isArray(json.value) ? json.value : null);
    if (!list || !list.length) { if (diag) diag.push('chat: no messages[]/value[] array'); return null; }
    const messages = [];
    let threadId, callStart;
    for (const m of list) {
      if (!m || typeof m !== 'object') continue;
      const conv = str(m.conversationid) || str(m.conversationId) || str(m.chatId) || (str(m.conversationLink) || '').split('/').pop();
      if (conv && THREAD.test(conv)) threadId = threadId || conv.match(THREAD)[0];
      const time = dt(m.composetime) || dt(m.originalarrivaltime) || dt(m.createdDateTime);
      const type = String(m.messagetype || m.messageType || '');
      if (/^Event\/Call|callStarted|systemEventMessage/i.test(type) || /<partlist|callStarted/i.test(String(m.content || ''))) {
        if (time && (!callStart || time < callStart)) callStart = time;
        continue;
      }
      if (!/^(Text|RichText(\/Html)?|message)$/i.test(type)) continue;
      const sender = str(m.imdisplayname) || (m.from && m.from.user && str(m.from.user.displayName)) || str(m.fromDisplayNameInToken) || '';
      const text = htmlToText(m.body && typeof m.body === 'object' ? m.body.content : m.content);
      if (!time || !text || m.deletetime || m.deletedDateTime) continue;
      messages.push({ id: String(m.id || m.clientmessageid || time + sender), time, sender, text });
    }
    if (!threadId || (!messages.length && !callStart)) { // only meeting chats
      if (diag) diag.push(`chat rejected: ${list.length} items, threadId ${threadId || 'none (no 19:meeting_ id)'}, kept ${messages.length}, types ${[...new Set(list.slice(0, 50).map((m) => m && (m.messagetype || m.messageType)))].join(',')}`);
      return null;
    }
    return { threadId, messages, callStart };
  }

  // ---- debug summary of one response (no secrets) ----
  function redactUrl(u) {
    try { const x = new URL(u); const q = [...x.searchParams.keys()]; return x.origin + x.pathname.replace(/[A-Za-z0-9_-]{60,}/g, '<long>') + (q.length ? ' ?params=' + q.join(',') : ''); }
    catch (e) { return String(u || '').split('?')[0].slice(0, 200); }
  }
  function redactVal(v) {
    if (typeof v !== 'string') return v;
    if (/^eyJ[A-Za-z0-9_-]{10,}\./.test(v) || /^(Bearer|Skype)/i.test(v) || /[A-Za-z0-9+/_-]{80,}/.test(v)) return '<redacted>';
    return v.length > 80 ? v.slice(0, 80) + '…' : v;
  }
  const DATE_KEY = /(date|time|start|end|created|modified|compose|arrival)/i;
  const SECRET_KEY = /(token|auth|secret|password|cookie|signature|tempauth|key$)/i;
  function interesting(k, v) {
    if (SECRET_KEY.test(k)) return false;
    if (typeof v === 'string') return isDate(v) || /19:[^\s"]*@thread/.test(v) || (DATE_KEY.test(k) && v.length < 60);
    if (typeof v === 'number') return (v > 1.5e12 && v < 2.5e12) || (DATE_KEY.test(k) && v > 1.5e9 && v < 2.5e9);
    return false;
  }
  function summarize(url, status, contentType, text) {
    const out = { url: redactUrl(url), status, contentType: String(contentType || '').slice(0, 60), size: text ? text.length : 0 };
    if (!text) { out.parse = 'skipped (empty or over 1 MB)'; return out; }
    let json;
    try { json = JSON.parse(text); } catch (e) { out.parse = 'not json'; return out; }
    out.topKeys = json && typeof json === 'object' ? Object.keys(json).slice(0, 30) : typeof json;
    const hits = [];
    (function walk(n, path, depth) {
      if (!n || typeof n !== 'object' || depth > 9 || hits.length >= 25) return;
      const keys = Array.isArray(n) ? n.slice(0, 5).map((_, i) => i) : Object.keys(n);
      for (const k of keys) {
        const v = n[k];
        if (v && typeof v === 'object') walk(v, path + '.' + k, depth + 1);
        else if (interesting(String(k), v)) hits.push([path.slice(1) + '.' + k, redactVal(v)].join(' = '));
        if (hits.length >= 25) break;
      }
    })(json, '', 0);
    out.hits = hits;
    const diag = [];
    try { out.meetingsFound = scanMeetings(json, diag).length; } catch (e) { diag.push('scanMeetings error: ' + e.message); }
    try { const c = scanChat(json, diag); out.chatFound = c ? c.messages.length : null; } catch (e) { diag.push('scanChat error: ' + e.message); }
    out.parserNotes = diag.slice(0, 8);
    return out;
  }

  const mightContain = (t) => typeof t === 'string' && /19:meeting_|meetingStartDateTime|startDateTime|MeetingStartTime|recordingStartDateTime/.test(t);
  const api = { ticks, scanMeetings, scanChat, mightContain, htmlToText, summarize, redactUrl };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else Object.defineProperty(root, '__teamsTranscriptMeta', { value: api, configurable: false });
})(typeof globalThis !== 'undefined' ? globalThis : this);
