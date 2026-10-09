const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { toReadable, detectExtension, parseOffset, formatTime } = require('../lib/convert.js');
const { baseName, sanitise } = require('../lib/naming.js');
const fx = (f) => fs.readFileSync(path.join(__dirname, 'fixtures', f), 'utf8');

test('VTT without speaker tags (like Teams .vtt downloads)', () => {
  const r = toReadable(fx('no-speakers.vtt'));
  assert.strictEqual(r.format, 'vtt');
  assert.strictEqual(r.hasSpeakers, false);
  assert.deepStrictEqual(r.lines, [
    '[00:03] Good morning, can everyone hear me?',
    '[00:06] Yes, loud and clear. Shall we start with the agenda?',
    '[01:15] Sure, first item is the test plan.',
  ]);
});

test('VTT with <v Name> tags merges consecutive speaker lines', () => {
  const r = toReadable(fx('with-speakers.vtt'));
  assert.strictEqual(r.hasSpeakers, true);
  assert.deepStrictEqual(r.lines, [
    "[00:01] Alice Example: Hello everyone. Let's get started.",
    '[00:05] Bob Sample: Thanks, Alice & welcome.',
    '[62:03] Alice Example: Wrapping up now.',
  ]);
});

test('Stream/Teams transcript JSON keeps speaker names', () => {
  const r = toReadable(fx('stream-transcript.json'));
  assert.strictEqual(r.format, 'json');
  assert.strictEqual(r.hasSpeakers, true);
  assert.deepStrictEqual(r.lines, [
    '[00:02] Example, Alice: Morning, quick round today. I finished the widget export.',
    "[01:10] Sample, Bob: Nice. I'll review it after lunch.",
  ]);
});

test('JSON with ISO 8601 offsets, out of order, CRLF VTT and BOM', () => {
  const r = toReadable(JSON.stringify({ entries: [
    { speakerDisplayName: 'B', text: 'second', startOffset: 'PT1M2.5S' },
    { speakerDisplayName: 'A', text: 'first', startOffset: 'PT3S' },
  ] }));
  assert.deepStrictEqual(r.lines, ['[00:03] A: first', '[01:02] B: second']);
  const crlf = '\uFEFFWEBVTT\r\n\r\n00:00:01.000 --> 00:00:02.000\r\n<v X>hi</v>\r\n';
  assert.deepStrictEqual(toReadable(crlf).lines, ['[00:01] X: hi']);
});

test('unparseable content returns null', () => {
  assert.strictEqual(toReadable('hello'), null);
  assert.strictEqual(toReadable('{"value":[]}'), null);
  assert.strictEqual(toReadable('{broken'), null);
  assert.strictEqual(toReadable('WEBVTT\n\n'), null);
});

test('helpers', () => {
  assert.strictEqual(parseOffset('00:01:02.4578205'), 62.4578205);
  assert.strictEqual(parseOffset('PT1H'), 3600);
  assert.strictEqual(formatTime(3725.9), '62:05');
  assert.strictEqual(detectExtension(fx('stream-transcript.json')), 'json');
  assert.strictEqual(detectExtension(fx('no-speakers.vtt')), 'vtt');
  assert.strictEqual(detectExtension('plain'), 'txt');
});

test('file names use the Berlin date and a safe meeting name', () => {
  assert.strictEqual(baseName({ name: 'Daily Standup.json', createdDateTime: '2026-10-08T22:30:00Z' }), '2026-10-09_Daily-Standup');
  assert.strictEqual(baseName({ name: 'Transcript.vtt', itemName: 'Weekly Sync-20261009_083000-Meeting Recording.mp4' }), '2026-10-09_Weekly-Sync');
  assert.strictEqual(baseName({ name: 'en-US', itemName: 'Q&A: plan/review?.mp4', createdDateTime: '2026-03-01T12:00:00Z' }), '2026-03-01_Q&A-plan-review');
  assert.strictEqual(sanitise('...'), 'transcript');
  assert.match(baseName({ name: 'X' }, new Date('2026-12-31T23:30:00Z')), /^2027-01-01_X$/);
});
