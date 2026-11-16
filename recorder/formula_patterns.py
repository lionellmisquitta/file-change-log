"""Deterministic, reversible A1 formula-fill patterns for saved snapshots.

A pattern is an observed *final state difference*, never proof of a fill command.
Both full before and after snapshots remain the authoritative exact evidence.
"""
import re
from openpyxl.formula.translate import Translator, TranslatorError

CELL = re.compile(r"^([A-Z]+)([1-9][0-9]*)$")


def pack_formula_runs(changes, max_patterns=200, min_run=3):
    """Return (patterns, ids_to_omit) only for homogeneous consecutive formula runs.

    Every formula is translated to the first cell of its run and must match.
    This supports relative and mixed ($) references without inventing commands.
    No pattern is emitted for discontiguous rows, singleton values, or translation
    failures. Restoring exact workbook contents requires the saved snapshots.
    """
    candidates = []
    for change in changes:
        formula = change.get("new_value")
        if not isinstance(formula, str) or not formula.startswith("="):
            continue
        m = CELL.fullmatch(change.get("address", ""))
        if m is None:
            continue
        candidates.append((change["sheet"], m.group(1), int(m.group(2)),
                           change["address"], formula, change))
    candidates.sort(key=lambda x: (x[0], x[1], x[2]))
    patterns, removed = [], set()
    group = []

    def flush():
        if len(group) < min_run or len(patterns) >= max_patterns:
            return
        first = group[0]
        patterns.append({
            "sheet": first[0], "column": first[1],
            "first_row": first[2], "last_row": group[-1][2],
            "cell_count": len(group), "anchor_address": first[3],
            "anchor_formula": first[4],
            "encoding": "relative_a1_translation",
            "observation": "saved_formula_states_not_proof_of_fill_command",
            "exact_prior_and_final_state_requires_snapshots": True,
        })
        removed.update(id(item[5]) for item in group)

    for item in candidates:
        if not group:
            group = [item]
            continue
        prev = group[-1]
        matching = item[0] == prev[0] and item[1] == prev[1] and item[2] == prev[2] + 1
        if matching:
            try:
                matching = Translator(item[4], origin=item[3]).translate_formula(group[0][3]) == group[0][4]
            except (TranslatorError, ValueError, TypeError):
                matching = False
        if matching:
            group.append(item)
        else:
            flush()
            group = [item]
    flush()
    return patterns, removed
