# Verification evidence

The cross-platform core tests and bundled Windows runtime import check passed in
[GitHub Actions](https://github.com/lionellmisquitta/file-change-log/actions/runs/40324113942).
The initial Windows failure exposed short-path alias handling; the fix normalises
existing real paths before checking scope. The failing regression test was retained.

Local browser checks exercised configuration, actual saved-file capture, explanations,
pause/resume, inert HTML in notes, and mobile/desktop layout. No JavaScript errors
remained and no horizontal overflow remained at 390px.

Actual desktop Excel callbacks and sign-in startup are NOT_VERIFIED. Hosted Windows
runners test Python and COM imports; they do not supply desktop Excel. Follow the
Windows acceptance checklist and record actual results before operational use.

`core-tests.txt` contains local command output; `report.json` identifies the tested
candidate. A subsequent documentation-only commit may add this evidence.
