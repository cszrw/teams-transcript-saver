const test = require('node:test');
const assert = require('node:assert');
const { scanForTranscripts, mightContainTranscript } = require('../lib/scan.js');
const metadata = require('./fixtures/item-metadata.json');

test('finds media.transcripts[] in a SharePoint item response', () => {
  const r = scanForTranscripts(metadata);
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].name, 'Weekly Sync.json');
  assert.strictEqual(r[0].language, 'en-US');
  assert.strictEqual(r[0].transcriptType, 'subtitle');
  assert.match(r[0].url, /streamContent\?tempauth=/);
  assert.strictEqual(r[0].itemName, 'Weekly Sync-20261009_083000-Meeting Recording.mp4');
  assert.strictEqual(r[0].createdDateTime, '2026-10-09T06:30:12Z');
});

test('finds nested transcript links in other shapes', () => {
  const r = scanForTranscripts({
    value: [{ name: 'Rec.mp4', createdDateTime: '2026-01-02T10:00:00Z',
      recap: { transcriptList: [{ fileName: 'Team Call.vtt', temporaryDownloadUrl: 'https://x.svc.ms/a?tempauth=1' }] } }],
  });
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].name, 'Team Call.vtt');
  assert.strictEqual(r[0].itemName, 'Rec.mp4');
});

test('accepts link by transcriptType or URL path, ignores unrelated download links', () => {
  assert.strictEqual(scanForTranscripts({ a: { transcriptType: 'subtitle', temporaryDownloadUrl: 'https://h.sharepoint.com/x' } }).length, 1);
  assert.strictEqual(scanForTranscripts({ a: { temporaryDownloadUrl: 'https://h.sharepoint.com/media/transcripts/1/streamContent' } }).length, 1);
  assert.strictEqual(scanForTranscripts({ thumbnails: [{ temporaryDownloadUrl: 'https://h.sharepoint.com/thumb.jpg' }] }).length, 0);
  assert.strictEqual(scanForTranscripts({ transcripts: [{ temporaryDownloadUrl: 'javascript:alert(1)' }] }).length, 0);
});

test('de-duplicates and handles junk', () => {
  const t = { transcriptType: 'subtitle', temporaryDownloadUrl: 'https://h.sharepoint.com/t' };
  assert.strictEqual(scanForTranscripts({ a: t, b: [t] }).length, 1);
  assert.deepStrictEqual(scanForTranscripts(null), []);
  assert.deepStrictEqual(scanForTranscripts('x'), []);
  const deep = {}; let n = deep; for (let i = 0; i < 50; i++) { n.transcripts = {}; n = n.transcripts; }
  assert.doesNotThrow(() => scanForTranscripts(deep));
});

test('pre-check', () => {
  assert.ok(mightContainTranscript(JSON.stringify(metadata)));
  assert.ok(!mightContainTranscript('{"value":[]}'));
});
