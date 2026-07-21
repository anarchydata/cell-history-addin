# Cell History – Excel Add-in

An Excel task-pane add-in that remembers where you've been in a workbook:

- **Back button** (on the Home tab of the ribbon, and in the pane) that jumps to your previous cell location. A **Forward** button in the pane lets you retrace your steps the other way.
- **History pane** listing your recent cell locations (newest first). Click any entry to jump straight to it, including across sheets.
- **Dwell timer** – a location is only recorded after your cursor stays there for **5 seconds** (configurable, 1–60s, in the pane). Rapidly clicking through cells doesn't pollute the history; staying on the same cell just refreshes that entry's timestamp instead of creating a duplicate.

The add-in uses the Office **shared runtime**, so the ribbon Back button and the tracking logic share the same state, and tracking keeps working while the pane is closed (after the add-in has loaded once in the session).

## Prerequisites

- Windows with Excel (Microsoft 365 desktop)
- [Node.js](https://nodejs.org/) 18+

## Run it

```powershell
cd cell-history-addin
npm install
npm run icons        # generates ribbon icon PNGs into assets/
npm start            # starts the dev server + sideloads the add-in into Excel
```

`npm start` handles the HTTPS dev certificate (you'll be prompted to trust it the first time), starts the local server on `https://localhost:3000`, and opens Excel with the add-in sideloaded.

When you're done:

```powershell
npm run stop
```

To run just the dev server without sideloading (e.g. if the add-in is already registered):

```powershell
npm run dev-server
```

## How the tracking works

1. Every selection change (including switching sheets) starts a 5-second countdown, shown live in the pane.
2. If you move again before the countdown finishes, the candidate is discarded and the countdown restarts at the new location.
3. If you stay put, the location is committed to history and becomes the "current" entry.
4. Re-visiting the current location (or sitting on it) updates its timestamp and visit count rather than adding a duplicate.
5. **Back** steps backward through committed locations; **Forward** steps forward again. Navigating via Back/Forward/clicking an entry does not itself create new history entries.

History (capped at 100 entries), the back/forward position, and the dwell-time setting are stored in the workbook via document settings, so they survive closing and reopening the file — save the workbook to persist them to disk. **Clear** empties the history.

## Project layout

| File | Purpose |
|------|---------|
| `manifest.xml` | Add-in manifest (shared runtime, ribbon buttons) |
| `src/taskpane/taskpane.html/.css/.js` | Task pane UI and all tracking/navigation logic |
| `server.js` | Minimal HTTPS static server using the Office dev certs |
| `assets/generate-icons.js` | Regenerates the ribbon icon PNGs |
