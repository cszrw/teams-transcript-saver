const test = require('node:test');
const assert = require('node:assert');
const { scanMeetings, scanChat } = require('../lib/meta.js');
const out = require('../lib/output.js');
const TID = '19:meeting_abcDEF123@thread.v2';

test('graph event / onlineMeeting gives start + thread', () => {
  const r = scanMeetings({ value: [{ subject: 'Daily Atlassian', start: { dateTime: '2026-10-19T07:30:00.0000000', timeZone: 'UTC' },
    end: { dateTime: '2026-10-19T08:15:00', timeZone: 'UTC' }, onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/' + encodeURIComponent(TID) + '/0' } }] });
  assert.strictEqual(r[0].start, '2026-10-19T07:30:00.000Z');
  assert.strictEqual(r[0].threadId, TID);
});

test('chatsvc messages', () => {
  const c = scanChat({ messages: [
    { id: '2', conversationid: TID, messagetype: 'RichText/Html', composetime: '2026-10-19T07:40:00Z', imdisplayname: 'Tolkmitt, Martin', content: '<p>link &amp; more</p>' },
    { id: '1', conversationid: TID, messagetype: 'Event/Call', composetime: '2026-10-19T07:31:00Z', content: '<partlist type="started"/>' },
    { id: '3', conversationid: TID, messagetype: 'ThreadActivity/AddMember', composetime: '2026-10-19T07:32:00Z', content: 'x' }] });
  assert.strictEqual(c.callStart, '2026-10-19T07:31:00.000Z');
  assert.deepStrictEqual(c.messages.map((m) => m.text), ['link & more']);
  const info = out.resolve({}, { meetings: [], chats: { [TID]: c } }, 'Daily Atlassian');
  assert.strictEqual(info.start, '2026-10-19T07:31:00.000Z');
  assert.match(out.header('Daily Atlassian', '2026-10-19', info), /Meeting start: 2026-10-19 09:31 Europe\/Berlin/);
  assert.match(out.chatSection(info), /\[2026-10-19 09:40\] Tolkmitt, Martin: link & more/);
  const j = JSON.parse(out.augmentJson('{"type":"Transcript","entries":[]}', info));
  assert.strictEqual(j.meetingStart, info.start); assert.strictEqual(j.chat.length, 1); assert.deepStrictEqual(j.entries, []);
});

test('no metadata: unchanged json body, unknown header', () => {
  const info = out.resolve({}, null, 'X');
  assert.strictEqual(info.start, undefined);
  assert.match(out.header('X', '2026-10-21', info), /unknown/);
});

test('debug summary redacts and finds dates/threads', () => {
  const { summarize } = require('../lib/meta.js');
  const r = summarize('https://x.sharepoint.com/a/b?tempauth=SECRET&x=1', 200, 'application/json',
    JSON.stringify({ accessToken: 'eyJabc.def.ghi', item: { createdDateTime: '2026-10-19T07:30:00Z', chat: '19:meeting_x@thread.v2' } }));
  assert.ok(!JSON.stringify(r).includes('SECRET') && !JSON.stringify(r).includes('eyJ'));
  assert.match(r.url, /params=tempauth,x/);
  assert.ok(r.hits.some((h) => h.startsWith('item.createdDateTime')));
  assert.ok(r.hits.some((h) => h.includes('19:meeting_x')));
});

test('real recap shapes (19 Oct): driveItem, collab ticks, substrate -> one meeting', () => {
  const T = '19:meeting_MDAwMDAwMDAtMDAwMC0wMDAwLTAwMDAtMDAwMDAwMDAwMDAw@thread.v2';
  const { scanMeetings, ticks } = require('../lib/meta.js');
  assert.strictEqual(ticks('639279937919530198').slice(0, 16), '2026-10-19T08:03');
  const drive = scanMeetings({ id: '01AM', source: { threadId: T }, media: { recordingStartDateTime: '2026-10-19T08:03:02.2982942Z', recordingEndDateTime: '2026-10-19T10:19:51.5782942Z' }, listItem: { createdDateTime: '2026-10-19T08:03:02Z' } });
  const collab = scanMeetings({ resources: [{ metadata: { startTime: '639279937919530198', endTime: '639280020012330198', threadId: T } }] });
  const search = scanMeetings({ EntitySets: [{ ResultSets: [{ Results: [{ Source: { MeetingStartTime: '2026-10-19T08:03:02Z', MeetingRecording: { TeamsThreadId: T } } }] }] }] });
  assert.strictEqual(drive.length, 1); assert.strictEqual(drive[0].source, 'recording'); assert.strictEqual(drive[0].threadId, T);
  assert.strictEqual(collab[0].source, 'collab'); assert.strictEqual(search[0].source, 'search');
  const meetings = [...collab, ...search, ...drive].map((x, i) => ({ ...x, seenAt: i }));
  const info = out.resolve({}, { meetings, chats: {} }, 'Introduction into Enterprise Architecture & requirements for CMDB/ITAM');
  assert.strictEqual(info.start, '2026-10-19T08:03:02.298Z');
  assert.match(info.startSource, /recording/);
  assert.match(out.header('X', '2026-10-19', info), /Meeting start: 2026-10-19 10:03 Europe\/Berlin \(end 12:19\)/);
  const naming = require('../lib/naming.js');
  assert.strictEqual(naming.baseName({ name: 'X.json', meetingDate: info.start }), '2026-10-19_X');
});
