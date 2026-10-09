# Teams Transcript Saver

A small Chrome extension that saves the **full** transcript of a Teams meeting from the
Teams web client. The Teams transcript pane only shows part of the transcript at a time, so
copying from it misses most of the meeting. This extension gets the whole file instead.

## What it does

1. When you open a meeting's Recap / Transcript in Teams on the web, Teams asks SharePoint for
   the recording's details (`…sharepoint.com/_api/v2.0/drives/<drive>/items/<item>`). That
   answer contains `media.transcripts[]`, each with a short-lived `temporaryDownloadUrl`.
2. The extension notices that answer in the page and remembers the link for that tab.
   A small **⬇ Save transcript** button appears bottom-right (drag it if it's in the way), and
   the toolbar icon shows how many transcripts it caught.
3. When you click Save, it downloads the transcript using your own logged-in session and saves:
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.json` (or `.vtt`), the raw file, unchanged
   - `Downloads/teams-transcripts/<YYYY-MM-DD>_<meeting-name>.txt`, a readable version:
     `[mm:ss] Speaker: text`, with consecutive lines from the same speaker merged.

   It asks for the JSON version first (`?format=json`), because that includes **speaker names**
   (`speakerDisplayName`). If JSON isn't offered it falls back to the plain stream, which is
   WebVTT and usually has no speaker names. The date is the meeting date (from the recording's
   name or creation time) in Europe/Berlin time, otherwise today.

## Privacy

- Runs only in your own browser, with your existing Teams login. No accounts, no tokens to set up.
- Nothing is sent anywhere: no servers, analytics or third parties. The only network requests
  it makes are the transcript downloads, to the same Microsoft address Teams itself uses.
- The temporary link is kept in Chrome's in-memory session storage for the tab only. It's
  wiped when you close the tab, reload the page or quit Chrome, and it's never written to the
  console or to disk. (The link expires on its own after a short time anyway.)
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

If you change or update the files, click the reload ↻ icon on the extension's card.

## Daily use

1. Open the meeting in Teams **on the web** (teams.microsoft.com or teams.cloud.microsoft), for
   example from the chat or calendar → **Recap** → **Transcript**.
2. Click **⬇ Save transcript** (bottom-right), or the toolbar icon → **Save** next to the transcript you want.
3. The files appear in `~/Downloads/teams-transcripts/`.

If several transcripts were caught on the tab, the floating button saves the most recent one.
Use the toolbar popup to choose another.

## Troubleshooting

- **"The download link has expired"**: the link Teams handed out is short-lived. Close and
  reopen the Transcript (or Recap) pane, then click Save again straight away.
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
- `lib/scan.js`: finds transcript links in API answers
- `lib/convert.js`: JSON / WebVTT → readable text
- `lib/naming.js`: file names
- `tests/`: unit tests with synthetic fixtures. Run `node --test tests/` (Node 18+, no installs needed).
