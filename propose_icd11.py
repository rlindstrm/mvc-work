#!/usr/bin/env python3
"""Propose Estonian translations for MVC value sets from the ICD-11 Estonian translation.

Matches each concept by its English name (column D) against the English labels of the
ICD-11 Estonian translation export (titles, fully specified names, synonyms and narrower
terms that have an Estonian label) and writes the proposal into two new columns after
the existing ones. Existing cells are never changed.

Usage:
    python propose_icd11.py ICD11_EXPORT.xlsx [WORKBOOK.xlsx ...] [--output output]

Without workbooks, every .xlsx in input/ is processed.
"""
import argparse
import collections
import copy
import re
import warnings
from pathlib import Path

import openpyxl

from translate import COL_TERM, header_row

PROPOSAL_HEADER = "et-EE proposal (ICD-11)"
BASIS_HEADER = "ICD-11 match (English label)"

# ICD-11 export columns (1-based): LabelEN, LabelET, Property
ICD_EN, ICD_ET, ICD_PROPERTY = 6, 7, 10
ICD_PROPERTIES = {"Title", "FullySpecifiedName", "Synonym", "NarrowerTerm"}

# ICD-11 uses British spelling, SNOMED CT American; both sides are folded before matching.
SPELLING = [("oesoph", "esoph"), ("haem", "hem"), ("caec", "cec"), ("foet", "fet"),
            ("tumour", "tumor"), ("anaem", "anem"), ("ae", "e")]
LATERALITY = {"left": "Vasak", "right": "Parem"}

warnings.filterwarnings("ignore", message="Workbook contains no default style")
warnings.filterwarnings("ignore", message="Data Validation extension is not supported")


def fold(text):
    text = re.sub(r"\s+", " ", text.replace(" ", " ")).strip().casefold()
    for british, american in SPELLING:
        text = text.replace(british, american)
    return text


def load_icd11(path):
    """Map folded English label -> (Estonian label, English label); most frequent Estonian label wins."""
    found = collections.defaultdict(collections.Counter)
    english = {}
    ws = openpyxl.load_workbook(path, read_only=True).worksheets[0]
    for row in ws.iter_rows(min_row=2, values_only=True):
        en, et = row[ICD_EN - 1], row[ICD_ET - 1]
        if row[ICD_PROPERTY - 1] in ICD_PROPERTIES and en and et and str(et).strip():
            key = fold(str(en))
            found[key][str(et).strip()] += 1
            english.setdefault(key, str(en).strip())
    return {key: (labels.most_common(1)[0][0], english[key]) for key, labels in found.items()}


def lower_first(text):
    # Keep acronyms and proper nouns ("MRT", "Vateri") as they are.
    return text[0].lower() + text[1:] if len(text) > 1 and text[1].islower() else text


def lookup(name, icd):
    """Estonian label for an English name, also composing "Left X"/"Right X" from X."""
    if fold(name) in icd:
        return icd[fold(name)]
    m = re.match(r"(?i)(left|right) (.+)$", name)
    if m and fold(m[2]) in icd:
        et, en = icd[fold(m[2])]
        return f"{LATERALITY[m[1].lower()]} {lower_first(et)}", f"{m[1].capitalize()} + {en}"
    return None


def propose(name, icd):
    """Return (proposal, basis) or None, following the "X, tervik" / "X, struktuur" convention."""
    hit = lookup(name, icd)
    if hit:
        return hit
    m = re.match(r"(?i)entire (.+)$", name)
    if m and (hit := lookup(m[1], icd)):
        return f"{hit[0]}, tervik", f"Entire + {hit[1]}"
    m = re.match(r"(?i)structure of (.+)$", name) or re.match(r"(?i)(.+) structure$", name)
    if m and (hit := lookup(m[1], icd)):
        return f"{hit[0]}, struktuur", f"Structure + {hit[1]}"
    return None


def add_proposals(ws, icd):
    counts = collections.Counter()
    start = header_row(ws)
    if start is None:
        return counts
    # First free column after the header row, so nothing already in the sheet is overwritten.
    col = max(c.column for c in ws[start] if c.value not in (None, "")) + 1
    template = ws.cell(start, COL_TERM + 1)
    for offset, title in enumerate([PROPOSAL_HEADER, BASIS_HEADER]):
        cell = ws.cell(start, col + offset, title)
        cell._style = copy.copy(template._style)
        ws.column_dimensions[openpyxl.utils.get_column_letter(col + offset)].width = 50
    for row in range(start + 1, ws.max_row + 1):
        name = ws.cell(row, COL_TERM).value
        if not name:
            continue
        result = propose(str(name).strip(), icd)
        counts["proposed" if result else "no match"] += 1
        style = ws.cell(row, COL_TERM + 1)._style
        for offset in range(2):
            cell = ws.cell(row, col + offset, result[offset] if result else None)
            cell._style = copy.copy(style)
    return counts


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    here = Path(__file__).resolve().parent
    parser.add_argument("icd11", type=Path, help="ICD-11 Estonian translation export (.xlsx)")
    parser.add_argument("workbooks", type=Path, nargs="*")
    parser.add_argument("--output", type=Path, default=here / "output")
    args = parser.parse_args()

    workbooks = args.workbooks or sorted(p for p in (here / "input").glob("*.xlsx") if not p.name.startswith("~$"))
    if not workbooks:
        raise SystemExit("No value set workbooks given or found in input/")

    icd = load_icd11(args.icd11)
    print(f"ICD-11 English labels with an Estonian translation: {len(icd)}\n")

    args.output.mkdir(parents=True, exist_ok=True)
    for path in workbooks:
        wb = openpyxl.load_workbook(path)
        print(path.name)
        for ws in wb.worksheets[1:]:
            counts = add_proposals(ws, icd)
            print(f"  {ws.title}: " + ", ".join(f"{k} {v}" for k, v in counts.items()))
        target = args.output / f"{path.stem}_ICD-11-proposals.xlsx"
        wb.save(target)
        print(f"  -> {target}\n")


if __name__ == "__main__":
    main()
