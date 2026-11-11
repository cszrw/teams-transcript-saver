// Picks the meeting start + chat for a transcript and formats the extra output.
(function (root) {
  'use strict';
  const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
  const berlin = (iso, withDate) => new Intl.DateTimeFormat('sv-SE', Object.assign({ timeZone: 'Europe/Berlin', hour: '2-digit', minute: '2-digit' },
    withDate ? { year: 'numeric', month: '2-digit', day: '2-digit' } : {})).format(new Date(iso));

  // meta: { meetings: [{start,end,subject,threadId,seenAt}], chats: {threadId: {messages, callStart, seenAt}} }
  // Returns { start, end, subject, threadId, startSource, chat }
  function resolve(t, meta, title) {
    meta = meta || { meetings: [], chats: {} };
    const chats = meta.chats || {};
    const want = norm(title);
    const matches = (meta.meetings || []).filter((m) => want && norm(m.subject) === want);
    // Prefer a match with a chat we have, then a start nearest the transcript's own date, then latest seen.
    const ref = t.meetingDate || t.createdDateTime;
    matches.sort((a, b) => (!!chats[b.threadId] - !!chats[a.threadId]) ||
      (ref ? Math.abs(new Date(a.start) - new Date(ref)) - Math.abs(new Date(b.start) - new Date(ref)) : 0) ||
      (b.seenAt || 0) - (a.seenAt || 0));
    let m = matches[0];
    if (!m) {
      // No title match (recording/collab/search objects carry no subject). If every dated candidate
      // on this tab belongs to one meeting thread, it is this recap's meeting.
      const dated = (meta.meetings || []).filter((x) => x.threadId);
      const ids = [...new Set(dated.map((x) => x.threadId))];
      if (ids.length === 1) {
        const rank = { recording: 0, collab: 1, search: 2, meeting: 3 };
        m = dated.sort((a, b) => (rank[a.source] ?? 9) - (rank[b.source] ?? 9) || (b.seenAt || 0) - (a.seenAt || 0))[0];
      }
    }
    let threadId = m && m.threadId;
    if (!threadId) {
      // Only one meeting chat seen on this tab: that's the one the recap belongs to.
      const ids = Object.keys(chats);
      if (ids.length === 1) threadId = ids[0];
    }
    if (!m && threadId) m = (meta.meetings || []).filter((x) => x.threadId === threadId).sort((a, b) => (b.seenAt || 0) - (a.seenAt || 0))[0];
    const chat = threadId && chats[threadId];
    let start, startSource;
    if (t.meetingDate) { start = t.meetingDate; startSource = 'recording metadata'; }
    else if (m && m.start) { start = m.start; startSource = { recording: 'recording driveItem (media.recordingStartDateTime)', collab: 'Teams collab object (metadata.startTime)', search: 'substrate search (MeetingStartTime)' }[m.source] || 'meeting details'; }
    else if (chat && chat.callStart) { start = chat.callStart; startSource = 'meeting chat (call started)'; }
    return { start, end: m && m.end, subject: (m && m.subject) || title, threadId, startSource,
      chat: chat && chat.messages.length ? chat : null };
  }

  function header(title, date, info) {
    let h = `${title} (${date})\n`;
    h += info.start ? `Meeting start: ${berlin(info.start, true)} Europe/Berlin${info.end ? ` (end ${berlin(info.end)})` : ''}\n`
      : 'Meeting start: unknown (date in file name is the download date)\n';
    return h + '\n';
  }

  function chatSection(info) {
    if (!info.chat) return '';
    return '\n== Meeting chat ==\n' + info.chat.messages.map((m) =>
      `[${berlin(m.time, true)}] ${m.sender ? m.sender + ': ' : ''}${m.text.replace(/\n/g, '\n    ')}`).join('\n') + '\n';
  }

  // Adds top-level meeting fields + chat to the Stream transcript JSON; leaves everything else as is.
  function augmentJson(raw, info) {
    let j;
    try { j = JSON.parse(String(raw).replace(/^\uFEFF/, '')); } catch (e) { return raw; }
    if (!j || typeof j !== 'object' || Array.isArray(j)) return raw;
    const extra = { meetingStart: info.start || null, meetingEnd: info.end || null, meetingSubject: info.subject || null,
      meetingStartSource: info.startSource || null, chatThreadId: info.threadId || null,
      chat: info.chat ? info.chat.messages.map(({ time, sender, text }) => ({ time, sender, text })) : [] };
    return JSON.stringify(Object.assign(extra, j));
  }

  const api = { resolve, header, chatSection, augmentJson };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TranscriptOutput = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
