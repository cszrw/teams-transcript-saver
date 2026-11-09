# Teams Transcript Saver

A small Chrome extension that saves the **full** transcript of a Teams meeting from the
Teams web client. The Teams transcript pane only shows part of the transcript at a time, so
copying from it misses most of the meeting. This extension gets the whole file instead.

## What it does

1. When you open a meeting's Recap / Transcript in Teams **on the web**
   (`teams.cloud.microsoft` or `teams.microsoft.com`), the extension watches the page for
   the recording's SharePoint identity (drive / item / transcript ids). Recap currently
   loads those from the Teams meeting-content API and from SharePoint
   (`/_api/v2.0/…/items/…`, `/_api_cached/v2.1/…/cdnmedia/transcripts`, Stream `g_fileInfo`).
   It does **not** depend on the short-lived `temporaryDownloadUrl` / `tempauth` link that
   Teams used to hand out — those now 401/403.
2. A small **⬇ Save transcript** button appears bottom-right (drag it if it's in the way), and
   the toolbar icon shows how many transcripts it caught.
3. When you click Save, it downloads the transcript with your logged-in SharePoint cookies from
   `/_api/v2.1/drives/…/items/…/media/transcripts/…/content?format=json` (cookie-auth, not a
   CDN token). If that 401s, it re-lists the transcripts and retries. It saves:
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.json` (or `.vtt`), the raw file
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.txt`, a readable version:
     `[mm:ss] Speaker: text`, with consecutive lines from the same speaker merged
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>_debug.json`, a privacy-safe
     sidecar (capture flags and notes; **no** tokens, URLs, tenant names, or transcript body)

   JSON is preferred because it includes **speaker names** (`speakerDisplayName`). If JSON
   isn't offered it falls back to WebVTT. The date is the meeting date (from metadata,
   the recording's name, or creation time) in Europe/Berlin time, otherwise today.

## Privacy

- Runs only in your own browser, with your existing Teams login. No accounts, no tokens to set up.
- Nothing is sent anywhere: no servers, analytics or third parties. The only network requests
  it makes are the transcript downloads, to the same Microsoft address Teams itself uses.
- Drive/item ids and cookie-auth content URLs are kept in Chrome's in-memory session storage
  for the tab only. They are wiped when you close the tab, navigate away, or quit Chrome, and
  they are never written to the console. The `_debug.json` sidecar stores flags (has drive
  item, url kind, …) rather than tokens or hostnames.
- Saved files go only to your Downloads folder.
- It saves transcripts of meetings you can already open, using your own access. Follow your
  company's rules on storing meeting transcripts outside Teams.

### Permissions and why

| Permission | Why |
| --- | --- |
| `https://teams.microsoft.com/*`, `https://teams.cloud.microsoft/*` | The Teams web client (old and new address). The page watcher and the button run here. |
| `https://*.sharepoint.com/*` | Where recordings and transcripts live (`<tenant>-my.sharepoint.com`). The recap player may run in a SharePoint frame, and this is the transcript download host. |
| `https://*.svc.ms/*` | Microsoft's media service hosts (Stream/OneDrive media, e.g. `*.mediap.svc.ms`). Included in case a transcript link points there instead of SharePoint. Only used to download a link Teams itself handed out. |
| `downloads` | To save the files into `Downloads/teams-transcripts/`. |
| `storage` | Session (memory-only) storage of the captured links per tab, so they survive Chrome putting the extension's background worker to sleep. |

## Install (once)

1. Open `chrome://extensions` in Chrome.
2. Turn on **Developer mode** (top right).
3. Click **Load unpacked** and pick this folder (`teams-transcript-saver`).
4. Optional: pin the extension (puzzle icon → pin) to see the badge count.

If you change or update the files, click the reload ↻ icon on the extension's card, then
**reload the Teams tab** so the page scripts pick up the new code. If Chrome shows a new
host permission, click **Allow**.

## Daily use

1. Open the meeting in Teams **on the web** (teams.microsoft.com or teams.cloud.microsoft), for
   example from the chat or calendar → **Recap** → **Transcript**.
2. Click **⬇ Save transcript** (bottom-right), or the toolbar icon → **Save** next to the transcript you want.
3. The files appear in `~/Downloads/teams-transcripts/`.

If several transcripts were caught on the tab, the floating button saves the most recent one.
Use the toolbar popup to choose another.

## Troubleshooting

- **Only a `_debug.json` file, no `.json` / `.txt`**: the capture was seen but the download
  failed. Open `_debug.json`: `result.error` `expired` means SharePoint still returned 401/403
  after a cookie-auth retry. Close Recap, open Recap → Transcript again, then Save. Confirm
  you are on `teams.cloud.microsoft` (or `teams.microsoft.com`) in **Chrome**, not the desktop app.
- **"The download link has expired"**: v1.2.0 retries with `/content` using your SharePoint
  cookies. If you still see this, reopen Recap → Transcript so a fresh drive/item id is
  captured, then Save straight away.
- **No button / popup says nothing captured**: reload the Teams tab after (re)loading the
  extension, then open Recap → Transcript. The watcher must be running in that tab.
- **The .txt has no speaker names**: Teams only offered the WebVTT version for that meeting.
  The raw file is still saved.
- **Date on the filename looks off**: the recording metadata sometimes has no
  `meetingStartDateTime`. The extension then uses the `yyyymmdd` stamp in the recording name,
  or the item's `createdDateTime`. That does not affect the transcript download.
- **Teams desktop app**: not supported. Use Teams in Chrome.

## Limits / not verified

- **Not tested against a live tenant in this repo.** Capture and URL rewrite are covered by
  unit tests with fictional Contoso fixtures (`node --test tests/`). Microsoft can change
  these internal APIs at any time.
- The transcript JSON format (`entries[]` with `speakerDisplayName`, `startOffset`, `text`, and
  schema `http://stream.office.com/schemas/transcript.json`) is based on public examples of
  Stream/Teams transcript downloads, not confirmed on this tenant. That `?format=json` returns
  this JSON is also from public reports. If it doesn't, the extension falls back to the VTT.
- Times in the .txt are `mm:ss` from the start of the recording (minutes go past 59 for long meetings).

## Files

- `manifest.json`: extension definition (Manifest V3)
- `inject.js`: runs in the page; watches fetch/XHR answers and does the page-side download
- `content.js`: bridge between page and extension; draws the floating button
- `background.js`: keeps captures per tab, downloads, converts and saves
- `popup.html`, `popup.js`: toolbar popup
- `lib/scan.js`: finds transcript ids/links in API answers and builds cookie-auth `/content` URLs
- `lib/convert.js`: JSON / WebVTT → readable text
- `lib/naming.js`: file names
- `tests/`: unit tests with synthetic fixtures. Run `node --test tests/*.test.js` (Node 18+, no installs needed).
