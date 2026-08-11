---
name: check-publications
description: Check ORCID, Google Scholar, and PubMed for Brian's recent works and report which are missing from the CV and biosketch, with paste-ready citations. Report-only — never edits documents. Use when asked to check for new publications, abstracts, or preprints.
---

# Check for new publications

**Report-only skill.** Never edit any document here — produce a report; Brian
approves additions, which then go through `/update-cv`.

Identifiers (also in `CLAUDE.md`): ORCID `0000-0002-2748-1444`, Google Scholar
user `GXofPAYAAAAJ`, PubMed query `Anderson BM` + radiation oncology/medical
physics context.

## 1. Gather from sources (no single point of failure)

Run all three; a source failing is a note in the report, not a reason to stop.

**Primary — ORCID public API** (reliable, structured):
```
GET https://pub.orcid.org/v3.0/0000-0002-2748-1444/works
Accept: application/json
```
Gives work summaries (title, type, year, journal, external IDs). For full detail
on an entry: `GET .../work/{put-code}`. Collect: title, DOI (external-id type
"doi"), journal, publication date, type (journal-article, conference-abstract,
preprint…).

**Cross-check — Google Scholar** (catches abstracts/preprints ORCID misses):
fetch `https://scholar.google.com/citations?user=GXofPAYAAAAJ&hl=en&sortby=pubdate&cstart=0&pagesize=100`.
Scholar aggressively blocks automation — if you get a CAPTCHA/429/empty page,
**skip it and state in the report that Scholar was unavailable**; do not retry
in a loop or attempt CAPTCHA workarounds.

**Fallback — PubMed E-utilities:**
```
GET https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi?db=pubmed&term=Anderson+BM[Author]&retmax=200&sort=pub+date&retmode=json
GET .../esummary.fcgi?db=pubmed&id=<ids>&retmode=json
```
"Anderson BM" is a common name — keep only hits whose journal/affiliation fits
(radiation oncology, medical physics, interventional radiology, medical
imaging/AI; affiliations: UCSD, UNC Chapel Hill, MD Anderson, Georgia Tech).
List borderline hits separately as "uncertain — please confirm" rather than
silently dropping or including them.

## 2. Merge and de-duplicate

- Match on **DOI** when both records have one (case-insensitive, strip
  `https://doi.org/` prefix).
- Otherwise match on normalized title (lowercase, strip punctuation/whitespace;
  treat ≥0.9 similarity or clear prefix match as the same work — titles get
  truncated differently per source).
- Keep the union of metadata (e.g. DOI from ORCID + venue from Scholar).

## 3. Compare against the documents

Extract current entries from:
- `Clincal_CV/BMAnderson CV.docx` — Publications section, all five subsections
  (Papers, Book Chapters, Oral Presentations, Poster Presentations, Abstracts).
- `Biosketch_BMA.docx` — the selected products under section C.

Use `python-docx` (installed; pandoc is not). Match found works to document
entries by title (documents have no DOIs). Remember journal-name quirks, e.g.
"Medical Physics" appears in the CV as "The International Journal of Medical
Physics Research and Practice".

## 4. Report

Output a short report with these parts:

1. **Missing from the CV** — each with a paste-ready citation in CV style
   (see CLAUDE.md: authors `Surname Initials`, **Anderson B.M** bold, *title*
   italic, journal, MM/YYYY) and the subsection it belongs in. Note: bold/italic
   shown in markdown maps to run formatting when later added.
2. **Missing from the biosketch** — flagged as FYI only: the biosketch lists
   *selected* products, so absence is a choice, not an error. Just list recent
   works Brian may want to consider swapping in.
3. **Source discrepancies** — works found in one source but not another
   (especially: on Scholar/PubMed but missing from ORCID → Brian should update
   his ORCID record; SciENcv builds the common-form biosketch from it).
4. **Source status** — which sources succeeded/failed/were skipped.
5. **Uncertain matches** — possible other-"Anderson BM" hits needing his eyes.

Do not edit any file. End by asking which items Brian wants added.
