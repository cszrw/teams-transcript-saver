// File naming: <YYYY-MM-DD>_<sanitised meeting name>.<ext>, date in Europe/Berlin.
(function (root) {
  'use strict';

  function berlinDate(d) {
    // en-CA formats as YYYY-MM-DD
    return new Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Berlin', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
  }

  function stripName(n) {
    return String(n || '')
      .replace(/\.(json|vtt|docx|txt|mp4)$/i, '')
      .replace(/[-_ ]*Meeting (Recording|Transcript)$/i, '')
      .replace(/[-_ ]*\d{8}_\d{6}(UTC)?/i, '')
      .trim();
  }

  // Use the transcript's own name ("Daily Standup.json"); fall back to the
  // drive item name when the transcript name is generic ("Transcript", "en-US").
  function meetingTitle(t) {
    const own = stripName(t.name);
    const generic = !own || /^(transcript|subtitle|captions?)$/i.test(own) || /^[a-z]{2}(-[a-z]{2,4})?$/i.test(own);
    return (generic && stripName(t.itemName)) || own || 'transcript';
  }

  function sanitise(name) {
    const s = String(name)
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
      .replace(/[.\s]+$/g, '')
      .replace(/^[.\s]+/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/ /g, '-')
      .slice(0, 80);
    return s || 'transcript';
  }

  // Meeting date: explicit meeting date, then a yyyymmdd stamp in the item name, then createdDateTime, then now.
  function meetingDate(t, now) {
    const tryDate = (v) => { const d = v ? new Date(v) : null; return d && !isNaN(d) ? d : null; };
    const d1 = tryDate(t.meetingDate);
    if (d1) return berlinDate(d1);
    const m = /(\d{4})(\d{2})(\d{2})_\d{6}/.exec(String(t.itemName || '') + ' ' + String(t.name || ''));
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    const d2 = tryDate(t.createdDateTime);
    if (d2) return berlinDate(d2);
    return berlinDate(now || new Date());
  }

  function baseName(t, now) {
    return `${meetingDate(t, now)}_${sanitise(meetingTitle(t))}`;
  }

  const api = { baseName, sanitise, meetingDate, meetingTitle, berlinDate };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TranscriptNaming = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
