const test = require('node:test');
const assert = require('node:assert');
const {
  scanForTranscripts,
  mightContainTranscript,
  parseSharePointApiUrl,
  parseRecapHref,
  downloadCandidates,
  discoverUrls,
  debugCapture,
  ticksToIso,
  buildTranscriptUrl,
} = require('../lib/scan.js');
const metadata = require('./fixtures/item-metadata.json');
const noTemp = require('./fixtures/item-v2-no-tempurl.json');
const collection = require('./fixtures/transcripts-collection.json');
const mcps = require('./fixtures/mcps-contents.json');

test('finds media.transcripts[] in a SharePoint item response', () => {
  const r = scanForTranscripts(metadata);
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].name, 'Weekly Sync.json');
  assert.strictEqual(r[0].language, 'en-US');
  assert.strictEqual(r[0].transcriptType, 'subtitle');
  assert.match(r[0].url, /\/media\/transcripts\/a1b2c3d4-0000-4000-8000-000000000001\/content$/);
  assert.strictEqual(r[0].contentUrl, r[0].url);
  assert.strictEqual(r[0].itemName, 'Weekly Sync-20261009_083000-Meeting Recording.mp4');
  assert.strictEqual(r[0].createdDateTime, '2026-10-09T06:30:12Z');
  assert.ok(!/[?&]tempauth=/i.test(r[0].url));
});

test('rewrites a tempauth streamContent capture to cookie-auth /content first', () => {
  const [t] = scanForTranscripts(metadata);
  const urls = downloadCandidates(t);
  assert.match(urls[0], /\/media\/transcripts\/a1b2c3d4-0000-4000-8000-000000000001\/content\?format=json$/);
  assert.ok(urls.every((u) => !/tempauth/i.test(u) || urls.indexOf(u) > 0));
  assert.ok(urls.every((u) => !/cdnmedia/i.test(u)));
  const dbg = debugCapture(t);
  assert.strictEqual(dbg.hasTranscriptId, true);
  assert.strictEqual(dbg.hasDriveItem, true);
  assert.strictEqual(dbg.urlKind, 'content');
  assert.ok(!JSON.stringify(dbg).includes('tempauth'));
  assert.ok(!JSON.stringify(dbg).includes('contoso'));
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
  assert.ok(mightContainTranscript(JSON.stringify(mcps)));
  assert.ok(!mightContainTranscript('{"value":[]}'));
});

test('builds cookie-auth URLs from metadata that has ids but no temporaryDownloadUrl', () => {
  const r = scanForTranscripts(noTemp, 'https://contoso-my.sharepoint.com/_api/v2.0/drives/b!exampleDriveId/items/01EXAMPLEITEMID0000000000000000000000');
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].transcriptId, 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee');
  assert.strictEqual(r[0].meetingDate, '2026-11-05T08:30:00Z');
  assert.match(r[0].url, /\/_api\/v2\.1\/drives\/b!exampleDriveId\/items\/01EXAMPLEITEMID0000000000000000000000\/media\/transcripts\/aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee\/content$/);
});

test('collection responses pick up drive/item ids from the request URL', () => {
  const src = 'https://contoso-my.sharepoint.com/personal/alex_contoso_com/_api_cached/v2.1/drives/b!exampleDriveId/items/01EXAMPLEITEMID0000000000000000000000/cdnmedia/transcripts';
  const parsed = parseSharePointApiUrl(src);
  assert.strictEqual(parsed.kind, 'list');
  assert.strictEqual(parsed.driveId, 'b!exampleDriveId');
  assert.strictEqual(parsed.sitePath, '/personal/alex_contoso_com');
  const r = scanForTranscripts(collection, src);
  assert.strictEqual(r.length, 1);
  assert.match(r[0].url, /\/_api\/v2\.1\/.*\/media\/transcripts\/aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee\/content$/);
  assert.ok(!/cdnmedia|_api_cached/.test(r[0].url));
});

test('meeting-content (mcps) resources become SharePoint /content URLs', () => {
  const r = scanForTranscripts(mcps);
  assert.strictEqual(r.length, 1);
  assert.strictEqual(r[0].name, 'Weekly Sync');
  assert.strictEqual(r[0].driveId, 'b!exampleDriveId');
  assert.match(r[0].url, /\/content$/);
  assert.ok(r[0].meetingDate, 'startTime ticks should become a meetingDate');
  assert.match(r[0].location, /\/content$/);
});

test('recap hash links expose drive and item ids', () => {
  const href = 'https://teams.cloud.microsoft/_#/l/meetingrecap?driveId=b!exampleDriveId&driveItemId=01EXAMPLEITEMID0000000000000000000000&sitePath='
    + encodeURIComponent('https://contoso-my.sharepoint.com/personal/alex_contoso_com');
  const recap = parseRecapHref(href);
  assert.ok(recap);
  assert.strictEqual(recap.driveId, 'b!exampleDriveId');
  assert.strictEqual(recap.itemId, '01EXAMPLEITEMID0000000000000000000000');
  assert.strictEqual(recap.host, 'contoso-my.sharepoint.com');
  assert.strictEqual(recap.sitePath, '/personal/alex_contoso_com');
  const list = buildTranscriptUrl(recap, 'list');
  assert.match(list, /\/_api\/v2\.1\/drives\/b!exampleDriveId\/items\/01EXAMPLEITEMID0000000000000000000000\/media\/transcripts$/);
});

test('ticksToIso converts .NET ticks used by meeting-content metadata', () => {
  const iso = '2026-11-05T08:30:00.000Z';
  const ticks = Date.parse(iso) * 10000 + 621355968000000000;
  assert.strictEqual(ticksToIso(ticks), iso);
  assert.strictEqual(ticksToIso(String(ticks)), iso);
  assert.strictEqual(ticksToIso(42), undefined);
});

test('downloadCandidates try content, streamContent, collab, and site-path variants; never the recording file', () => {
  const t = {
    driveId: 'b!exampleDriveId',
    itemId: '01EXAMPLEITEMID0000000000000000000000',
    transcriptId: 'aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee',
    host: 'contoso-my.sharepoint.com',
    origin: 'https://contoso-my.sharepoint.com',
    sitePath: '/personal/alex_contoso_com',
    location: 'https://teams.cloud.microsoft/api/mcps/v1/contents/transcript',
  };
  const urls = downloadCandidates(t);
  assert.match(urls[0], /\/personal\/alex_contoso_com\/_api\/v2\.1\/.*\/content\?format=json$/);
  assert.ok(urls.some((u) => /\/streamContent\?format=json$/.test(u)));
  assert.ok(urls.some((u) => /applyhighlights=false/.test(u) && /applymediaedits=false/.test(u)));
  assert.ok(urls.some((u) => /^https:\/\/contoso-my\.sharepoint\.com\/_api\/v2\.1\//.test(u)));
  assert.ok(urls.some((u) => u === 'https://teams.cloud.microsoft/api/mcps/v1/contents/transcript'));
  assert.ok(urls.every((u) => !/cdnmedia|_api_cached/.test(u)));
  assert.ok(urls.every((u) => !/\/items\/[^/]+\/content(?:\?|$)/.test(u)), 'must not download the recording item');
  const disc = discoverUrls(t);
  assert.ok(disc.some((u) => /\/media\/transcripts$/.test(u)));
  assert.ok(disc.some((u) => /\/versions\/current\/media\/transcripts$/.test(u)));
  assert.ok(disc.some((u) => /\/media\/transcripts\/aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee$/.test(u)));
  assert.ok(disc.some((u) => /teams\.cloud\.microsoft\/api\/mcps/.test(u)));
});
