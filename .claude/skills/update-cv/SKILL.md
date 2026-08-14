---
name: update-cv
description: Add a new entry (publication, abstract, invited talk, committee role, grant, mentee, award, clinical project) to Brian's clinical CV, preserving Word formatting, and re-export the PDF. Use whenever the CV needs a new line item or an existing entry updated.
---

# Update the clinical CV

Target: `Clincal_CV/BMAnderson CV.docx` (source of truth). After any edit,
re-export `Clincal_CV/BMAnderson CV.pdf` — the pair must never drift.

Read `CLAUDE.md` first for full structure/citation conventions. Summary of where
things go:

| New item | Section | Format |
|---|---|---|
| Peer-reviewed paper | Publications → Papers | Numbered citation (style below) |
| Book chapter | Publications → Book Chapters | *Book title*, chapter, authors, publisher info |
| Conference talk (Brian presenting) | Publications → Oral Presentations (Presenting Author) | Citation + conference + city/"Virtual" + MM/YYYY |
| Conference poster (Brian presenting) | Publications → Poster Presentations (Presenting Author) | Same as oral |
| Abstract (co-author presenting) | Publications → Abstracts | Same, presenting author first |
| Invited talk / seminar / keynote | Invited Talks | `Org, role/series: "Title" (MM/YYYY)` — date always at the end |
| Grant / fellowship | Grants and Fellowships | Date column + underlined funder/award lead-in, then role, amount, period, description |
| Committee / editorial / society role | Professional Service Activities | Under the org's subhead — date column + underlined `Role:` lead-in. Mirror the change in `Professional Service Activities.docx` |
| Award | Honors and Awards | Date column + underlined `Org, Award name:` lead-in + description |
| Clinical tool/project | Clinical Projects and Responsibilities | Under the institution subhead — date column + underlined `Project:` lead-in (no site prefix) |
| Research position/project | Research Experience | Under the institution subhead — date column if dated (four legacy entries, UCSD ×2 / UNC ×2, are undated plain paragraphs) |
| Mentee | No dedicated section exists yet — ask Brian where he wants it (likely a new small-caps `Heading 1` section, e.g. "Mentorship") before inventing one |

## Date-column entry layout (every section except Publications and Invited Talks)

Dated entries use a left date column: `YYYY` / `YYYY – YYYY` / `YYYY – Present`
(occasionally `YYYY, YYYY`), then a tab, then the underlined lead-in and the
description after a colon. Paragraph properties: `Normal` style, left indent
2016 twips (1.4"), hanging first-line indent, left tab stop at 2016 twips —
copying a sibling entry's paragraph XML carries all of this over. Publications
and Invited Talks are the exception: auto-numbered lists with the date at the
**end** of each entry.

## Citation style (match exactly)

`Anderson B.M` **bold** wherever it appears in the author list (exactly that
form — no comma, no trailing period unless it ends the author list); title in
*italics* ending with a period; journal/conference plain, full journal names
(not abbreviations); date `MM/YYYY` at the end. Authors as `Surname Initials`
separated by commas. Example run breakdown:

> **Anderson B.M**, Moore L., Bojechko C. *Prediction of in-vivo Electronic Portal Imaging Device Transit Images with a Convolutional Neural Network…* Medical Physics 11/2024

## Ordering

Every dated list is reverse chronological — insert the new entry at the
**top** of its list (or in correct date position if backfilling older work).

## Editing rules — preserve formatting

1. Never rewrite or re-style paragraphs you aren't adding/changing. The doc is
   entirely Times New Roman; sections use styles `Heading 1` (small-caps,
   bottom border), `Normal`, and `List Paragraph` with auto-numbering.
2. Preferred method: unzip the .docx and edit `word/document.xml` directly —
   **copy an existing sibling paragraph's XML** (same style, numbering `numPr`,
   run properties) and change only the text runs. This guarantees the new entry
   inherits the list numbering and fonts. Re-zip with the same structure.
   Alternative: `python-docx` (installed), using
   `copy.deepcopy` of a sibling paragraph element, then edit its runs.
3. Bold/italic live on runs: Brian's name is its own bold run; the title is its
   own italic run. Build the new entry from 3+ runs, don't apply bold/italic to
   whole paragraphs.
4. Publications subsection headers ("Papers ____…") are Normal-style paragraphs
   padded with underscores — leave them alone.

## After editing

1. **"Last updated" date:** the CV currently has no such field. If Brian has
   added one by the time you read this (check the header/footer and the name
   block), update it to today's date on every edit.
2. **Re-export the PDF** with Word COM (faithful rendering — do not use
   LibreOffice unless Word fails):

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("C:\Users\Markb\Modular_Projects\CV_and_Biosketch\Clincal_CV\BMAnderson CV.docx")
$doc.SaveAs([ref]"C:\Users\Markb\Modular_Projects\CV_and_Biosketch\Clincal_CV\BMAnderson CV.pdf", [ref]17)  # 17 = wdFormatPDF
$doc.Close($false)
$word.Quit()
```

   If the .docx is open in Word (file-lock error), ask Brian to close it first.
3. Sanity-check the export: page count similar to before, new entry visible
   (extract text with `pypdf` and grep for the new title).
4. Offer to run `/sync-biosketch` if the change affects positions, honors, or
   could displace a "selected product."
5. Commit .docx + .pdf together with a message describing the added entry.
