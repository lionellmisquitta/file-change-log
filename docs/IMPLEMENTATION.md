# File Change Log

A local Windows recorder for selected folders and live Excel changes. Set it up
once, keep working in Excel, and review the events and snapshots later.

**Status: first pilot.** File capture, comparisons, cache behaviour and local HTTP
boundaries are tested outside Windows. Actual Excel COM events and Windows sign-in
startup must pass the [Windows acceptance check](docs/WINDOWS-ACCEPTANCE.md).
This is an observational recorder, **not an exhaustive or tamper-proof audit trail**.

## Start here

Use the **FileChangeLog-Windows-x64.zip** download when available in the development
PR's Actions artifacts. It bundles Python and dependencies; no separate Python or
pip installation is required. It targets Windows 10/11 x64, not native ARM64.

1. Extract into a permanent local folder. Double-click **Start Recorder.bat**.
2. Enter the folders to watch and click **Save folders**.
3. Click **Enable start at sign-in**. The recorder starts for your Windows user.
4. Leave it running and work normally. Use **Open Recorder.bat** to view it again.
5. Add explanations beside important changes. Export JSONL for later LLM review.

Normal Windows/app-control policies still apply. A portable package is not a way
to bypass your organisation's controls. Source checkout users need Python 3.11+
and `python -m pip install -r requirements.txt`, then `python run.py`.

## What it records

| Input | Live capture | Saved-file capture |
| --- | --- | --- |
| Desktop Excel `.xlsx`, `.xlsm`, `.xlsb`, `.xls` | Cell/range edits, current values/formulas; prior values where a baseline was observed; workbook save/open requests/results; sheet activation/create/delete requests; calculation and PivotTable-update notifications | Snapshots of saved bytes. Additional cell-coordinate comparison for `.xlsx`/`.xlsm` |
| CSV/TSV opened in Excel | Excel events if the desktop application exposes them | Snapshots and text comparison |
| Other text/data files | File-system notifications | Snapshots and bounded text comparison for supported extensions |
| Documents/PDFs/presentations | File-system notifications, not application-internal edits | Original snapshots; no detailed content comparison in this pilot |
| Other file types | File-system notifications and periodic state observations | No snapshots, matching the chosen scope |

Mac Excel, Excel for the web, Google Sheets, and edits in other users' sessions do
not generate live Excel events in this recorder. Net saved-file changes on shared
drives can still be observed. The recorder does not identify a remote cell editor.

## How it works

One Python process runs in the signed-in desktop session:

- `watchdog` listens for folder notifications. A periodic rescan reconciles saved
  state, including after missed notifications or network interruptions.
- `pywin32` attaches to desktop Excel using COM application events. It discovers
  the active application and registered workbook instances. Visible Excel processes
  that cannot be attached are reported as a capture gap, not silently counted as covered.
- A baseline cache provides previously observed cell values. The event itself
  supplies only the changed range; unknown prior values are marked unknown.
- Saved files are copied after their size/modified signature settles. A hash
  identifies the captured bytes. Originals are never rewritten.
- JSONL logs are written and flushed to disk; SQLite indexes them for the UI and
  stores restart state. The UI runs on loopback with a random session token.

There is no Gmail integration, cloud service, LLM call, automatic rule inference,
or ADW writeback. A service inbox may collect source mail separately.

## Where history lives

Default: `%LOCALAPPDATA%\FileChangeLog`

- `logs/events-YYYY-MM-DD.jsonl`: append-only event records, UTC timestamps.
- `snapshots/<path-hash>/<content-hash>.<extension>`: captured source versions.
- `index.sqlite`: internal, rebuildable event index and persistent file inventory.
- `settings.json`: selected folders and capture limits.

The application directory contains code only. **Do not commit user history to Git.**
Backup closed snapshot files and JSONL exports if desired. Do not sync the active
SQLite database through Drive. No history is automatically deleted.

Paths can be local, Google Drive for desktop paths, mapped network drives, or UNC.
Mapped drives are normalised to UNC when Windows can resolve the mapping. Cloud
placeholders may need to download before a snapshot succeeds. Subfolders are included;
symlinks/junctions are skipped. A monitored root cannot contain the recorder history.
Folder notifications depend on the filesystem/provider; scans are a fallback and
cannot reconstruct every transient file or intermediate version.

## Limits that are explicit in the logs

- Default maximum snapshot size: **512 MiB per file**, adjustable in the UI. Larger
  files generate events and an explicit snapshot-skipped record. Disk-space checks
  can also defer captures. No unlimited-storage assumption is hidden in the recorder.
- Default live baseline: **200,000 cells per sheet**. A larger used range gets an
  incomplete-baseline event. New values can still be captured within the event limit.
- Default live range limit: **10,000 cells per event**. Larger edits are logged as
  range-level events with omitted details and an invalidated old-value baseline.
- Snapshot comparison scans up to **200,000 cells**, with at most 5,000 changed-cell
  details per result. These comparisons are coordinate based: sorting or inserting
  cells can produce many differences and does not prove a business operation.
- Structural edits can invalidate coordinate assumptions. Do not treat cached old
  values as independent proof of row identity after insert/delete/move operations.
- Text comparison: files up to **2 MiB**; at most 5,000 diff lines. Original bytes
  are retained even where content comparison is unavailable or bounded.
- Recalculation and PivotTable events record that an update occurred, not every
  changed calculated result. Some notifications are throttled to one per sheet
  per 10 seconds to avoid flooding.
- Formatting/filtering/chart actions, complete clipboard provenance, exact
  undo/redo commands, and the reason for an edit are not exhaustive live events.
- Pauses, disconnects, disabled Excel events and incomplete scans create explicit
  gap/error records. There is no recording while the PC is asleep or shut down.
- A workbook-close request is not proof it closed; a save request is not proof it
  saved. Live edits remain separate from captured saved-file versions.
- The UI shows 200 recent matching events and exports the latest 1,000 matching
  events. The complete history remains in `logs/`.

## Build and verify

```sh
python -m unittest discover -s tests -v
python build_windows_bundle.py
```

The bundle builder downloads the pinned official CPython x64 embeddable runtime
and pinned PyPI wheels, verifies SHA256 hashes and includes package metadata/licenses.
Build output goes to `dist/` and is excluded from source control.

Read [CAPTURE-CONTRACT.md](docs/CAPTURE-CONTRACT.md) for the event schema and
[WINDOWS-ACCEPTANCE.md](docs/WINDOWS-ACCEPTANCE.md) before live use.

## Stop or remove

Use **Disable start at sign-in**, then **Stop recorder completely**. The startup
entry is `HKCU\Software\Microsoft\Windows\CurrentVersion\Run\FileChangeLog`.
If the application cannot launch, that named entry can be removed manually.
Files already recorded stay in the history directory until you delete them.
