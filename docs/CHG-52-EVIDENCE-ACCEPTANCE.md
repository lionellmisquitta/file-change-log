# CHG-52 — Exact saved-byte lineage and targeted explanation prompts

## Scope
1. When a **newly arrived** monitored file's saved snapshot matches the retained snapshot of another tracked path in the same watched root, emit `file_identical_content_observed`. Include the original saved snapshot reference, matching prior paths (maximum 20), and the capturing event ID.
2. Do **not** assert a user copy, Save As, formula transformation or other causal action. Identical bytes establish equality of saved versions only.
3. Show exact-content lineage in the timeline with the event reference.
4. In **Explanation Mode only**, prioritize optional prompts for observed formula overrides/large edit events bearing `explanation_wanted` and requested sheet deletions. Permit the user to dismiss or explain; Silent Mode must not prompt.
5. Preserve prior, less targeted suggestions as fallback.

## Acceptance
- Baseline file `a.csv` and later identical `b.csv`: explicit `file_identical_content_observed`, same snapshot, source path and event ID.
- Different saved bytes: no identical-content event.
- Existing worksheet/snapshot evidence, chronology, notes and exports remain intact.
- Targeted prompt remains opt-in and does not invent a human explanation or classify intent.
- Full Python tests + Windows package CI pass. Windows Excel event handling needs actual Windows pilot checklist `docs/WINDOWS-PIVOT-ACCEPTANCE.md`.

## Explicitly outside scope
- Cell-fingerprint probable copies (can cause false positives)
- Unsaved copy operations and complete provenance across moved folders and monitored roots
- Guaranteed capture of every structural Excel command
- Finished process replay or rule generation

**Stacked dependency:** CHG-52 is based on CHG-51, which is based on CHG-50. Merge in dependency order and re-run tests on resulting main. Source-only changes do not prove end-to-end Windows operation.
