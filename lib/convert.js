// Turns a Teams/Stream transcript (JSON or WEBVTT) into readable text:
//   [mm:ss] Speaker: text
// Used by background.js (via importScripts) and the node tests.
(function (root) {
  'use strict';

  // "00:01:02.4578205", "01:02.5", "PT1H2M3.5S", 62.5 (seconds) -> seconds
  function parseOffset(v) {
    if (typeof v === 'number' && isFinite(v)) return v;
    if (typeof v !== 'string') return NaN;
    const s = v.trim();
    const iso = /^P(?:T)?(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?$/i.exec(s);
    if (iso && s.length > 2) return (+iso[1] || 0) * 3600 + (+iso[2] || 0) * 60 + (+iso[3] || 0);
    const parts = s.replace(',', '.').split(':');
    if (parts.length < 2 || parts.length > 3 || parts.some((p) => p === '' || isNaN(+p))) {
      return /^\d+(\.\d+)?$/.test(s) ? +s : NaN;
    }
    return parts.reduce((acc, p) => acc * 60 + +p, 0);
  }

  function formatTime(sec) {
    if (!isFinite(sec)) return '--:--';
    const t = Math.floor(sec);
    return String(Math.floor(t / 60)).padStart(2, '0') + ':' + String(t % 60).padStart(2, '0');
  }

  function cleanText(t) {
    return String(t == null ? '' : t)
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // ---- JSON ----
  function findEntries(json) {
    if (Array.isArray(json)) return json;
    if (!json || typeof json !== 'object') return null;
    for (const k of ['entries', 'transcriptEntries', 'segments', 'items', 'phrases']) {
      if (Array.isArray(json[k])) return json[k];
    }
    if (json.transcript && typeof json.transcript === 'object') return findEntries(json.transcript);
    return null;
  }

  function parseJson(text) {
    let json;
    try { json = JSON.parse(text.replace(/^\uFEFF/, '')); } catch (e) { return null; }
    const entries = findEntries(json);
    if (!entries) return null;
    const cues = [];
    for (const e of entries) {
      if (!e || typeof e !== 'object') continue;
      const body = cleanText(e.text != null ? e.text : e.content != null ? e.content : e.displayText);
      if (!body) continue;
      cues.push({
        start: parseOffset(e.startOffset != null ? e.startOffset : e.start != null ? e.start : e.offset),
        speaker: cleanText(e.speakerDisplayName || e.speakerName || e.displayName ||
          (typeof e.speaker === 'string' ? e.speaker : e.speaker && e.speaker.displayName) || ''),
        text: body,
      });
    }
    return cues.length ? cues : null;
  }

  // ---- WEBVTT ----
  function parseVtt(text) {
    const src = text.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');
    if (!/^WEBVTT/.test(src)) return null;
    const cues = [];
    for (const block of src.split(/\n{2,}/)) {
      const lines = block.split('\n');
      const ti = lines.findIndex((l) => l.includes('-->'));
      if (ti === -1) continue; // header, NOTE, STYLE
      const start = parseOffset(lines[ti].split('-->')[0]);
      let speaker = '';
      const parts = [];
      for (const raw of lines.slice(ti + 1)) {
        const m = /<v(?:\.[^\s>]+)*\s+([^>]+)>/.exec(raw);
        if (m && !speaker) speaker = m[1].trim();
        const t = cleanText(raw);
        if (t) parts.push(t);
      }
      if (parts.length) cues.push({ start, speaker, text: parts.join(' ') });
    }
    return cues.length ? cues : null;
  }

  // Merge consecutive cues by the same (known) speaker; keep the first timestamp.
  function mergeCues(cues) {
    const out = [];
    for (const c of cues) {
      const prev = out[out.length - 1];
      if (prev && c.speaker && prev.speaker === c.speaker) prev.text += ' ' + c.text;
      else out.push({ ...c });
    }
    return out;
  }

  // Returns { format: 'json'|'vtt', lines: [...], hasSpeakers } or null if not parseable.
  function toReadable(text) {
    if (typeof text !== 'string') return null;
    const trimmed = text.replace(/^\uFEFF/, '').trimStart();
    let format = null;
    let cues = null;
    if (trimmed[0] === '{' || trimmed[0] === '[') { cues = parseJson(trimmed); format = 'json'; }
    else if (/^WEBVTT/.test(trimmed)) { cues = parseVtt(trimmed); format = 'vtt'; }
    if (!cues) return null;
    const hasStart = cues.every((c) => isFinite(c.start));
    if (hasStart) cues = cues.map((c, i) => ({ ...c, i })).sort((a, b) => a.start - b.start || a.i - b.i);
    const merged = mergeCues(cues);
    return {
      format,
      hasSpeakers: merged.some((c) => c.speaker),
      lines: merged.map((c) => `[${formatTime(c.start)}] ${c.speaker ? c.speaker + ': ' : ''}${c.text}`),
    };
  }

  // Guess a file extension for the raw content.
  function detectExtension(text) {
    const t = String(text || '').replace(/^\uFEFF/, '').trimStart();
    if (/^WEBVTT/.test(t)) return 'vtt';
    if (t[0] === '{' || t[0] === '[') { try { JSON.parse(t); return 'json'; } catch (e) { /* fall through */ } }
    return 'txt';
  }

  const api = { toReadable, detectExtension, parseOffset, formatTime, parseVtt, parseJson, mergeCues };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.TranscriptConvert = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
