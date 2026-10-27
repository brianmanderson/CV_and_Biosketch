# NSF Collaborators and Other Affiliations (COA) — Brian M. Anderson

Document (3) of the four NSF senior/key personnel documents (PAPPG 24-1, Chapter II.D.2.h).
Everything in this folder is a **proposal for Brian to review**; nothing has been submitted or committed.

## Files

| File | What it is |
|---|---|
| `coa_template_original.xlsx` | Pristine copy of NSF's COA template, downloaded 2026-10-27 from `https://nsf-gov-resources.nsf.gov/files/coa_template.xlsx` (26,889 bytes, MD5 `7c5f07e816f56d99598756d77522716a`, server Last-Modified 2025-02-10). Never edit. |
| `Anderson_BM_NSF_COA.xlsx` | The filled COA. Rows were added with openpyxl inside Table 4's row area only; headers, columns, sheet name, page setup, data validations and the five Excel table objects are intact. Upload this .xlsx to Research.gov as-is (Research.gov converts it). |
| `COA_source_list.md` | Traceability: every Table 4 name with the CV entry (title, date) that qualifies it, the organization entered, the source URL for that organization, the reason for every `[CONFIRM]` flag, per-table row counts, CV transcription issues, and the open questions. |
| `tooling/` | The scripts and data used to fill the template (see *Regenerating*), plus the seven raw affiliation-lookup tables in `tooling/lookups/`. |
| `README.md` | This file. |

## Reference date and look-back windows

**Assumed submission date: 10/27/2026.** All windows below are computed from that date.
If the proposal is submitted on a different date, recompute the windows and re-check the
entries nearest each boundary (the two 10/2022 papers; see open question 6 in `COA_source_list.md`).

| Table | NSF rule | Window used |
|---|---|---|
| 1 — Your affiliations | Organizational affiliations in the last 12 months | 10/27/2025 – 10/27/2026 |
| 2 — Personal, family, business relationships | No time limit stated by NSF | — |
| 3 — Advisors/advisees | G (Ph.D. advisor), T (Ph.D. thesis advisee): no time limit | — |
| 4 — Co-authors (A) and collaborators (C) | Last 48 months (publication may be later) | 10/27/2022 – 10/27/2026; publications dated 10/2022 or later are included |
| 5 — Editorial board / co-editors | Last 24 months | 10/27/2024 – 10/27/2026 |

Entries are not sorted beyond alphabetical order in Table 4. "Last Active" is optional and left blank
everywhere (all relationships are treated as ongoing). Each person is listed with their **current**
organization, not where the collaboration happened.

## What was entered

| Table | Rows | Content |
|---|---|---|
| 1 | 2 | University of California, San Diego (primary); University of North Carolina at Chapel Hill (Adjunct Assistant Professor, 2025–present, per the NIH biosketch) |
| 2 | 0 | Left empty — only Brian can answer this |
| 3 | 1 | G: Kristy K. Brock, The University of Texas MD Anderson Cancer Center (Ph.D. advisor 2017–2021). No Ph.D. thesis advisees appear in the CV |
| 4 | 175 | All type "A:" (co-author). 173 people from the CV's Publications section dated 10/2022 or later (30 entries), plus 2 co-authors (K. Yu, C. T. Lin) who appear on the ASTRO 2026 listing of one abstract but not in the CV entry. 37 rows carry `[CONFIRM]` in the Optional column; the full reason for each is in `COA_source_list.md`. No type "C:" rows — every named project collaborator is already a co-author in the window |
| 5 | 1 | B: John M. Boone, University of California, Davis — Editor-in-Chief of *Medical Physics* (2021–2026; Brian is an Associate Editor 2023–present). Co-editors Brian worked with directly are unknown and flagged |

Sources: `Clincal_CV/BMAnderson CV.docx` (extracted with python-docx), `Biosketch_BMA.docx` (UNC adjunct appointment),
and the web pages cited per row in `COA_source_list.md` (faculty directories, 2025–2026 PubMed/Europe PMC bylines,
ORCID, conference listings; lookups done 2026-10-27).

## Fidelity notes on the filled .xlsx

- The openpyxl-written file was post-processed (`tooling/postprocess.py`) to put back parts openpyxl rewrites:
  the five Excel table definitions (with their accessibility alt-text and Excel revision ids, ranges updated),
  the original column definitions (hidden columns keep width 0), the inactive `sheetProtection` element, and
  the page header/footer (openpyxl re-encodes the header's newline as literal `_x000a_`, which printed at the top of every page).
- Still dropped relative to the original, none of which affect content or layout: three SharePoint
  `customXml` metadata parts, the binary printer-settings part, and a Mac-Excel page-layout-view extension.
  Shared strings are written as inline strings.
- Verified: cell-by-cell comparison of values and styles against the pristine template outside the filled
  areas (0 differences), openpyxl load/save round trip, Excel (COM) open with all five table objects and
  validations in place, LibreOffice render to a 5-page PDF.

## Regenerating

All scripts are in `tooling/` and were run from a scratch folder with `python -I` (openpyxl 3.1.5 installed).

1. `roster.py` — hand-normalized author lists for the 30 qualifying CV entries, with spelling-variant merges.
2. `assemble.py <out.md>` — merges `lookups/batch*.md` with the roster, applies the name/organization corrections
   listed at the top of the script, writes `table_data.json` and `COA_source_list.md`.
3. `fill_coa.py <template.xlsx> <table_data.json> <out.xlsx>` — locates each table by its header labels, clears the
   sample rows, inserts rows inside Table 4 (shifting data validations, Excel table ranges and row heights by hand),
   writes the values, saves.
4. `postprocess.py <template.xlsx> <openpyxl_out.xlsx> <final.xlsx>` — restores the parts described above.
5. `verify_coa.py <template.xlsx> <final.xlsx> <table_data.json>` — the comparison and content checks.
