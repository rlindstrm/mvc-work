# MVC translation

Adds Estonian (et-EE) translations to MyHealth@EU Master Value Set Catalogue (MVC) value sets.

```
input/    value set workbooks (.xlsx) exported from the MVC
sources/  Estonian SNOMED CT extension release (SnomedCT_*.zip, or the extracted folder)
output/   translated workbooks
```

## Usage

```
pip install -r requirements.txt
python translate.py
```

The script prints how many concepts were translated per value set and writes one
workbook to `output/` for each input workbook.

## What the script does

For every workbook in `input/`:

- The first sheet (catalogue info) is copied unchanged.
- On every other sheet, the concept table is located by its `Code System ID` header row.
  For each row whose code system is SNOMED CT (`2.16.840.1.113883.6.96`), column E is
  replaced with the Estonian **preferred term**: an active Estonian description that is
  an active member of the Estonian language reference set (`71000181105`) with
  acceptability *Preferred* (`900000000000548007`). If there is none, column E gets `N/A`.
- Rows from other code systems are left unchanged.
- Columns B, D and E are set to fixed widths so the long header texts do not stretch them.
- The output file is named after the value sets it contains, e.g.
  `MVC_eHDSIProcedures_et-EE.xlsx`.

If `sources/` holds more than one SNOMED CT release, the newest (by release date in the
file name) is used.

## Data files are not committed

This repository is public, and SNOMED CT release files may only be distributed under
licence, so `.gitignore` keeps the contents of `input/`, `sources/` and `output/` out of git.
