# File Change Log

Record activity in selected Windows folders and detailed desktop Excel events, keep saved snapshots, and add explanations for later review.

**First pilot:** automated Windows/Linux tests and Windows package checks pass. Actual desktop Excel event capture and automatic sign-in startup still need the target-machine acceptance check. This is an observational tool with documented gaps, not a complete audit trail.

## Try it

1. [Open the validated Windows build](https://github.com/lionellmisquitta/file-change-log/actions/runs/40324466170).
2. Sign into GitHub, scroll to **Artifacts**, and download **FileChangeLog-Windows-x64**. Extract the artifact and then the application ZIP inside it if needed.
3. Extract the application into a permanent local folder and double-click **Start Recorder.bat**.
4. Add the folder paths and select **Save folders**.
5. Select **Enable start at sign-in** for automatic startup at future logins.

Python and dependencies are included in the Windows x64 package. Company application-control policies may still require an allow-list decision. **Code → Download ZIP downloads source code, not the bundled Windows application.**

## Demonstrate it

Use [the five minute demo guide](https://github.com/lionellmisquitta/file-change-log/blob/CHG-001/windows-recorder/docs/DEMO.md). The meeting package includes a synthetic browser walkthrough and a helper that creates disposable sample Excel files. The browser walkthrough is a replay, not live monitoring.

## What it captures

- File creation, modification, rename and deletion notifications, with periodic reconciliation scans.
- Live desktop Excel cell/range changes and selected workbook/sheet events where Excel exposes them.
- Old values where a baseline was observed; explicit unknowns and limits otherwise.
- Saved snapshots of spreadsheets, documents and text/data files.
- Structured JSONL logs and explanations for later analysis.

It supports selected local, mapped/shared and Google Drive for desktop folders. Other file types receive file-event logging without snapshots. It never writes into the monitored source files. The current pilot has no mailbox ingestion, cloud upload, LLM processing or ADW integration.

## Restarts and everyday use

After **Enable start at sign-in**, the recorder reloads the selected folders and history automatically when that Windows user signs in. It reconnects to Excel when available and rescans saved files. Keep the extracted app in the same location.

| Situation | Behaviour |
| --- | --- |
| Windows restart followed by sign-in | Starts quietly if startup was enabled |
| Browser tab closed | Recording continues |
| Sleep and resume | Monitoring resumes; intermediate events cannot be reconstructed |
| Shared drive unavailable | Logs a gap and retries; does not infer mass deletion |
| Process crash or manual stop | Reopen Start Recorder.bat or wait for next sign-in; no crash watchdog yet |

History defaults to `%LOCALAPPDATA%\FileChangeLog`. No history is automatically deleted. Saved snapshots and exported logs can be backed up; do not sync the active SQLite database. A log of an unsaved edit is not proof the workbook was saved.

## Share with a colleague

Share this public repository link for the overview. For installation, attach the downloaded Windows application ZIP or use your approved shared-drive link. GitHub Actions downloads require GitHub sign-in and have retention limits. Do not share a `sandbox:` link copied from a ChatGPT conversation as the installation link.

Actual workbooks, captured logs, snapshots and business-specific plans stay outside this public repository.

## Source and technical detail

- [Implementation branch](https://github.com/lionellmisquitta/file-change-log/tree/CHG-001/windows-recorder)
- [Development PR](https://github.com/lionellmisquitta/file-change-log/pull/1)
- [Implementation guide](https://github.com/lionellmisquitta/file-change-log/blob/CHG-001/windows-recorder/docs/IMPLEMENTATION.md)
- [Capture contract](https://github.com/lionellmisquitta/file-change-log/blob/CHG-001/windows-recorder/docs/CAPTURE-CONTRACT.md)
- [Windows acceptance checklist](https://github.com/lionellmisquitta/file-change-log/blob/CHG-001/windows-recorder/docs/WINDOWS-ACCEPTANCE.md)

The implementation is on the development branch and has not been merged into main. The main branch is the sharing overview.
