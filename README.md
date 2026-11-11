# Teams Transcript Saver

A small Chrome extension that saves the **full** transcript of a Teams meeting from the
Teams web client. The Teams transcript pane only shows part of the transcript at a time, so
copying from it misses most of the meeting. This extension gets the whole file instead.

It does **not** use the Recap player's native **Download** button. That control can be
disabled even when you can still open the transcript; the extension talks to the same
cookie-authenticated SharePoint transcript APIs Teams already uses.

## What it does

1. When you open a meeting's Recap / Transcript in Teams on the web, Teams talks to
   SharePoint (`…sharepoint.com/_api/…/drives/<drive>/items/<item>/…`). Recap often uses a
   short-lived `temporaryDownloadUrl` or a `_api_cached/…/cdnmedia/transcripts` request
   that 401/403s. The extension records the **drive, item and transcript ids** from those
   calls (and from recap links / meeting-content JSON) instead of depending on the token.
2. A small **⬇ Save transcript** button appears bottom-right (drag it if it's in the way),
   and the toolbar icon shows how many transcripts it caught.
3. When you click Save, it downloads the transcript with your logged-in browser session:
   cookie-auth `…/_api/v2.1/drives/…/items/…/media/transcripts/<id>/content?format=json`.
   If `/content` returns 401/403 it retries `streamContent`, Teams collab/mcps endpoints,
   and the item's transcripts collection. It does not download the recording file.
   Files saved:
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.json` (or `.vtt`), the raw file
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.txt`, a readable version:
     `[mm:ss] Speaker: text`, with consecutive lines from the same speaker merged.

   It asks for the JSON version first (`?format=json`), because that includes **speaker names**
   (`speakerDisplayName`). If JSON isn't offered it falls back to the plain stream, which is
   WebVTT and usually has no speaker names. The date is the meeting date (from meeting
   details, recording metadata, or the recording's name) in Europe/Berlin time, otherwise today.

## Privacy

- Runs only in your own browser, with your existing Teams login. No accounts, no tokens to set up.
- Nothing is sent anywhere: no servers, analytics or third parties. The only network requests
  it makes are the transcript downloads, to the same Microsoft address Teams itself uses.
- Captured ids and links are kept in Chrome's in-memory session storage for the tab only. They
  are wiped when you close the tab, reload the page or quit Chrome, and they are never written
  to the console. Debug dumps omit download URLs and tokens.
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
**reload the Teams tab**.

## Daily use

1. Open the meeting in Teams **on the web** (teams.microsoft.com or teams.cloud.microsoft), for
   example from the chat or calendar → **Recap** → **Transcript**.
2. Click **⬇ Save transcript** (bottom-right), or the toolbar icon → **Save** next to the transcript you want.
3. The files appear in `~/Downloads/teams-transcripts/`.

If several transcripts were caught on the tab, the floating button saves the most recent one.
Use the toolbar popup to choose another.

The native Teams **Download** button is not required and may be disabled.

## Troubleshooting

- **"The download link has expired"**: 1.2.0 retries cookie-auth `/content`, `streamContent`,
  collab, and the transcripts collection. If it still fails, close and reopen the Transcript
  (or Recap) pane, then click Save again. Reload the Teams tab after reloading the extension.
- **No button / popup says nothing captured**: the extension has to be running *before* Teams
  asks for the transcript. Close and reopen the transcript pane, or reload the Teams page
  (Cmd+R) and open the transcript again. If you just installed or reloaded the extension,
  reload the Teams tab.
- **The .txt has no speaker names**: Teams only offered the WebVTT version for that meeting.
  The raw file is still saved.
- **Teams desktop app**: not supported. Use Teams in Chrome.

## Limits / not verified

- **Not tested against the live Teams page.** It was built from the request seen in DevTools
  and tested against a local mock of that SharePoint response in headless Chrome
  (capture from fetch and XHR, same-origin and cross-origin, badge, floating button, JSON save,
  VTT fallback, expired-link message). Microsoft can change these internal APIs at any time.
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
- `lib/scan.js`: finds transcript ids and builds cookie-auth download URLs
- `lib/convert.js`: JSON / WebVTT → readable text
- `lib/naming.js`: file names
- `lib/meta.js`: meeting date and chat-thread resolution from page JSON
- `lib/output.js`: .txt header, chat section, JSON augment
- `icons/`: toolbar icons
- `tests/`: unit tests with synthetic **contoso** fixtures. Run `node --test tests/*.test.js`
  (Node 18+, no installs needed).

## 1.1.0: meeting date and chat

- The file date is now the meeting's start date (Europe/Berlin), taken from the meeting details or the
  meeting chat that the Teams page loads while you have it open. If neither was seen, it falls back to the
  download date and the .txt says `Meeting start: unknown`.
- `.txt` line 2 is now `Meeting start: YYYY-MM-DD HH:MM Europe/Berlin (end HH:MM)`. If chat messages were
  seen, the file ends with a `== Meeting chat ==` section: `[YYYY-MM-DD HH:MM] Sender: text`.
- `.json` gets extra top-level fields: `meetingStart`, `meetingEnd`, `meetingSubject`, `meetingStartSource`,
  `chatThreadId`, `chat` (`[{time, sender, text}]`, times in UTC ISO). `entries`/`events` are unchanged.
- To get the chat: open the meeting's **Chat** tab (scroll up to load older messages) before saving.

## 1.1.x (kept in 1.2.0)

- Timeouts on every save step so the button cannot sit on "Saving…" forever.
- Button watchdog if Teams removes the floating control from the DOM.
- Guard against "Extension context invalidated" after an extension reload (reload the Teams tab).
- Overwrite previous files in `teams-transcripts/` instead of creating ` (1)` copies.
- Meeting-date and chat-thread resolution from recap/collab JSON (`lib/meta.js`).

## 1.2.0: cookie-auth transcript download

- Save no longer uses the Recap `temporaryDownloadUrl` / `_api_cached/cdnmedia` stream as the
  first choice. Those expire and 401/403; a first 401 used to abort with an empty transcript
  body (name / language / capturedAt only).
- Downloads with the page's cookies at
  `/_api/v2.1/drives/{drive}/items/{item}/media/transcripts/{id}/content?format=json`.
- On 401/403, refreshes the listing and also tries `streamContent` (including
  `applyhighlights=false&applymediaedits=false`), Teams collab/mcps locations, and the item's
  `media/transcripts` collection. Does not require the recording **Download** permission.
- Captures drive/item/transcript ids from current Recap requests, recap URLs, and
  meeting-content JSON even when the response is binary cdnmedia.
