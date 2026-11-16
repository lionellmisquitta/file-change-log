# File Change Log

**See what changed. Preserve the evidence. Understand the process.**

![File Change Log animated walkthrough](media/file-change-log-demo.gif)

**[Watch the complete LinkedIn walkthrough (MP4)](media/file-change-log-linkedin.mp4)** · A compact, repository-friendly copy of the supplied video.


File Change Log is a local-first Windows application that observes changes in selected folders and Excel workbooks, retains saved file versions, and lets people add explanations to the resulting activity history.

It is designed for workflows where understanding *how* a spreadsheet or collection of files evolved matters: financial and royalty operations, recurring reporting, reconciliation, data preparation, and process discovery.

**Current status:** `v0.2.4-beta` code on `main` · Windows x64 · Open-source repository · **Beta, not a complete audit trail**

[**Get the Windows app**](https://github.com/lionellmisquitta/file-change-log/actions/workflows/verify.yml) · [**Releases**](https://github.com/lionellmisquitta/file-change-log/releases) · [**Report an issue**](https://github.com/lionellmisquitta/file-change-log/issues) · [**How capture works**](docs/CAPTURE-CONTRACT.md)

> **Important:** Saved-workbook pivot extraction and automated pivot-operation comparisons have passed tests. End-to-end *live Microsoft Excel* COM recording, Excel slicer reconstruction, and complete process replay have **not** yet passed target-desktop acceptance testing.

## Download and install

### Windows x64 — no separate Python installation

1. Open the [**latest Windows build workflow**](https://github.com/lionellmisquitta/file-change-log/actions/workflows/verify.yml). Select the newest **successful** run and scroll down to **Artifacts**.
2. Download **`FileChangeLog-Windows-x64`**. GitHub may require you to sign in.
3. Extract the downloaded artifact. If it contains `FileChangeLog-Windows-x64.zip`, extract that ZIP too.
4. Move the extracted application folder to a **permanent local directory**, for example `C:\Users\your-name\Apps\FileChangeLog`. **Do not run from within a ZIP.**
5. Double-click **`Start Recorder.bat`**. The dashboard opens in your browser.
6. Add a full folder path per line under **Folders to watch** and click **Save folders**. Subfolders are included.
7. Optionally select **Enable start at sign-in** to restart the recorder when your Windows user signs in.

The packaged app includes its own Python runtime. Admin rights are not normally required, but an organization-managed PC may require IT to allow the bundled runtime or BAT launcher.

**First-run check:** Open a disposable workbook inside a monitored folder, edit a cell, and save it. Verify `file_version_captured`; if the Excel listener is attached, check `excel_range_changed` as well. A saved-file event alone does **not** prove unsaved Excel edits are being captured.

### Upgrade without losing history

Stop the existing recorder from its dashboard, extract the new package into a permanent folder, and launch `Start Recorder.bat`. Verify monitored folder settings and startup configuration after upgrading.

**Never delete `%LOCALAPPDATA%\FileChangeLog` during an upgrade.** This is the default location for your settings, JSONL event history, SQLite index and retained snapshots; it is separate from the application folder. Back up that directory before changing installations. Do not copy an active SQLite database while the recorder is writing to it.

### Reopen, pause or stop

Closing the browser tab **does not stop recording**. Run `Open Recorder.bat` to reopen the dashboard. Use **Pause recording** for a documented gap, or **Stop recorder completely** to stop the process.

**Security note:** If Windows Smart App Control or company policy blocks the launcher/runtime, contact your IT administrator for an approved allow-list process. Do not disable device protections to run the app.

**Download note:** GitHub Actions artifacts expire and may require sign-in. A GitHub **Code → Download ZIP** is source code, **not** the packaged Windows app. Once a permanent release asset is published, use the [Releases page](https://github.com/lionellmisquitta/file-change-log/releases) for that version.

## Why this exists

Many operational processes are carried out through spreadsheets and files without a formal workflow engine. A final workbook shows the result, but usually does not explain which formulas were overwritten, which file version preceded it, how a pivot configuration evolved, or why someone made a particular change.

File Change Log builds an **evidence trail**, not an invented narrative: observed changes and retained file versions first, deterministic comparisons second, human explanations when useful.

## What it does today

| Capability | What is captured |
| --- | --- |
| **Folder activity** | Created, modified, moved and missing-on-scan files; filesystem notifications plus reconciliation scans |
| **Saved versions** | Content-addressed snapshots of supported files, bounded by configurable size/storage limits |
| **Live Excel observation** | Selected workbook, worksheet and changed-cell/range events where desktop Excel COM supplies them and listener attachment succeeds |
| **Cell differences** | Saved XLSX/XLSM cell-coordinate comparisons, including formula changes and flagged formula-to-value overrides |
| **Formula patterns** | Compact records for repeated relative/mixed-reference formula states, with individual exceptions preserved |
| **Pivot tables** | Saved pivot definitions and changes to source, rows, columns, values, aggregations and page-field arrangement |
| **Filters** | Supported worksheet and Excel Table saved filter metadata; unsupported structures explicitly flagged |
| **Activity timeline** | Chronological recent observations with raw event IDs, workbook paths, snapshot links and known gaps |
| **Optional explanations** | Notes linked to recorded events; Silent Mode does not interrupt work |
| **Evidence export** | ZIP with raw history and checksums; retained snapshots included when explicitly selected |

**What the recorder does not do:** It does not automatically upload your workbooks or logs to a cloud service or send them to an AI model. It does not modify monitored source files. It is not an infallible audit log and cannot infer every user action from a before/after comparison.

## How it works

```text
Selected folders and desktop Excel
             |
             +-- File-system notifications and periodic scans
             |     -> saved snapshots (content-addressed)
             |     -> saved-file cell/pivot/filter comparisons
             |
             +-- Excel COM event listener (Windows desktop)
                   -> observed workbook, sheet, range and pivot-update events
                              |
                              v
                    Local JSONL event history
                     + SQLite index
                              |
                  Timeline, notes and evidence ZIP
```

The recorder distinguishes **observed facts** from interpretations. For example, when a pivot first appears in a saved workbook, it records `first_observed`, not an unproven claim that a user clicked *Create PivotTable*. Identical saved bytes likewise do not prove a copy operation.

History defaults to **`%LOCALAPPDATA%\FileChangeLog`**. Snapshot capture is bounded: exceeding a limit can leave metadata observations while preserving an explicit capture gap.

## Pivot-table recording: what's been verified

The automated tests include both a public Excel workbook containing genuine native pivot definitions and a real spreadsheet-engine operation sequence.

- **5/5** genuine pivot definitions discovered from a public workbook, with axis-field counts checked against the underlying XML.
- Automated creation, Sum → Count, column-field addition, page-field addition, and removal across **five saved-state transitions**.
- **Six workbook versions** retained in the automated operation sequence; snapshot bytes and evidence-ZIP checksums verified.
- A separate **1,000-transaction** public-data comparison accounted for all **1,001** tested cell changes, with repeated formula patterns and manual exceptions preserved.

The interactive operation sequence was performed by **LibreOffice Calc through UNO**, then recorded as saved XLSX versions. That is meaningful saved-state integration coverage, **not** proof of Microsoft Excel's live COM event ordering or of complete pivot/slicer recreation.

[Real pivot automated QA](docs/PIVOT-OPERATIONS-AUTOMATED-QA.md) · [Public-data QA](tests/REAL_PIVOT_PUBLIC_DATA_QA.md) · [Windows Excel acceptance checklist](docs/WINDOWS-PIVOT-ACCEPTANCE.md)

## Current beta limitations

- **Live Microsoft Excel COM:** Actual desktop Excel listener attachment and unsaved pivot activity still require end-to-end testing on a Windows machine with Microsoft Excel installed.
- **Slicers and advanced pivot behavior:** Slicer selections/relationships, some grouped fields, refresh causes and complete pivot replay are not yet reconstructed.
- **Coverage gaps:** Large changes, inaccessible files, pauses, locked network folders, sleep, process interruptions and snapshot limits can make a recording incomplete. Known gaps are logged rather than fabricated.
- **Snapshot semantics:** A saved workbook comparison shows changed state, **not necessarily the command or person responsible**.
- **Platform:** Windows x64 is the target application platform; Linux is used in automated test infrastructure.

Use representative copies while evaluating. Do not rely on the beta as the sole regulatory or legal audit record.

## Test and develop

The source application is Python-based. On a development machine with Python and the project requirements installed:

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
```

GitHub Actions runs Windows and Ubuntu regression suites and packages the Windows application. A separate workflow runs real pivot-operation tests through headless LibreOffice Calc and its UNO interface.

[Regression and Windows packaging](https://github.com/lionellmisquitta/file-change-log/actions/workflows/verify.yml) · [Real pivot operation tests](https://github.com/lionellmisquitta/file-change-log/actions/workflows/real-pivot-operations.yml)

## Project direction

The longer-term goal is to move from **evidence collection → deterministic workbook logic extraction → explainable workflow reconstruction**. Automated replay, comprehensive pivot/slicer reconstruction and process inference are **future work**, not features of this beta.

## Feedback and contributions

Have a reproducible problem, missing event, or workflow you'd like this tool to understand? [Open a GitHub Issue](https://github.com/lionellmisquitta/file-change-log/issues) and include the Windows version, tested application build, expected/actual result, and sanitized event details.

**Do not attach confidential workbooks, customer data, passwords or raw business activity logs to public issues.**

---

**File Change Log** — local-first operational evidence for file and Excel workflows. Built for transparency about *what was observed, what was missed, and what remains unknown*.
