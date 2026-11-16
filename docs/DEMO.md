# Five minute demonstration

## Before the meeting

Use copies and synthetic data. Extract the Windows x64 package into a permanent
local folder and launch Start Recorder.bat. Run Create Demo Files.bat in the
meeting package. It creates a new uniquely named folder under Documents and
never overwrites an existing demo folder.

Add that folder in the recorder. Wait for the initial snapshot, then open its
Demo Orders.xlsx in desktop Excel. Confirm the status reports an attached Excel
instance before describing live edit capture as demonstrated.

## Demonstration sequence

1. Explain: We are observing repetitive preparation work before deciding what to automate.
2. Show the selected folder and initial saved-file snapshot.
3. Change cell C2 from 100 to 10. Show the live range-change record, if attached.
4. Save the workbook. Show the separate saved-file capture and comparison.
5. Click Explain and add: Synthetic example only. Quantity corrected from users
   to systems for this period. This is not an approved reusable rule.
6. Export the log. Explain that an LLM may later draft rules from these records,
   but the pilot itself makes no LLM calls and does not change the source files.

## If the live environment is blocked

Open demo/Walkthrough.html in a browser. It is explicitly labelled as a synthetic
replay and records nothing. Use it to explain the intended interaction; do not
present it as proof that live Excel integration worked.

## Sharing

- Public overview: https://github.com/lionellmisquitta/file-change-log
- Implementation branch: https://github.com/lionellmisquitta/file-change-log/tree/CHG-001/windows-recorder
- Current development PR: https://github.com/lionellmisquitta/file-change-log/pull/1
- Validated build: https://github.com/lionellmisquitta/file-change-log/actions/runs/40324466170

GitHub Actions artifact downloads require sign-in and expire. For a meeting,
attach the downloaded Windows ZIP or use an approved shared-drive file link.
The repository overview is public. Do not upload actual source documents,
captured histories, vendor contracts, or stakeholder-specific implementation plans.
