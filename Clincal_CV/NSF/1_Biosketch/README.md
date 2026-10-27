# NSF Biographical Sketch (SciENcv) — working folder

Prepared 27 October 2026. Everything here is a proposal for Brian to review; nothing has been entered into SciENcv and no existing document (CV, NIH biosketch, PDFs) was changed.

## Files

| File | What it is |
|---|---|
| `NSF_Biosketch_SciENcv_Worksheet.docx` | Field-by-field entry sheet mirroring the SciENcv NSF Biographical Sketch form, with the source of every value. Yellow `[CONFIRM]` tags mark values the source documents disagree on or omit. |
| `NSF_Biosketch_SciENcv_Worksheet.pdf` | PDF export of the worksheet (Word). Re-export after any .docx change. |
| `build_worksheet.py` | Regenerates the .docx with python-docx. Edit the data lists at the top of the file, run `python build_worksheet.py`, then re-export the PDF. |

## What NSF requires

- PAPPG NSF 24-1 (still current on 27 Oct 2026, with supplements NSF 26-200 and 26-202), Chapter II.D.2.h(i): the Biographical Sketch for every senior/key person **must be generated in SciENcv** using the NSPM-33 common form (OMB 3145-0279). A hand-made PDF is not accepted.
- Sections: Identifying Information; Organization and Location; Professional Preparation; Appointments and Positions; Products (up to 5 "most closely related" + up to 5 "other significant"); Certification. No page limit; no Synergistic Activities section (that is a separate one-page upload).
- Appointments: every **titled** academic, professional or institutional position, paid or not, including adjunct/visiting/honorary. Appointments outside the primary organization within the three years before submission must be listed — so both UNC appointments stay in.
- Certification must be done by Brian in his own SciENcv account, dated within 12 months of submission.
- The certified PDF already in the repo (`..\..\Anderson_CommonFormBiosketch.pdf`, 26 Feb 2026) is the **NIH** common form with the NIH supplement. It cannot be submitted to NSF; a new NSF document is needed.

## SciENcv click path

1. Open <https://www.ncbi.nlm.nih.gov/sciencv/> and click **Log in** (top right). Sign in with **eRA Commons** (user name `BrianMAnderson`), **ORCID**, Login.gov, or the NSF/Research.gov account — use whichever login is already linked to the NCBI account that holds the February 2026 biosketch, so the existing profile data is reused.
2. From the NCBI dashboard open **SciENcv**, then click **Create New Document**.
3. In the dialog: give the document a name (for example `NSF Biosketch 2026 <proposal>`), choose the format **NSF Biographical Sketch**, and pick the data source:
   - **External Source → ORCID** imports education, employment and works from ORCID (see the ORCID warning below), or
   - **Existing Document** copies one of Brian's SciENcv documents (if the February 2026 NIH biosketch is offered, this carries over the professional-preparation start dates and the appointments), or
   - **Start with a blank document** and type everything from the worksheet.
4. Fill the sections in worksheet order. All fields are required unless marked optional; the state and country fields are drop-downs.
   - **Identifying Information / Organization and Location** — worksheet section 1.
   - **Professional Preparation** — worksheet section 2, newest first by start date (residency, Ph.D., S.M.S., B.S.).
   - **Appointments and Positions** — worksheet section 3a, newest first by start date; section 3b lists the optional roles.
5. **Products**: click the add/select-citations control in each of the two lists. SciENcv offers citations from **My Bibliography**, which can be filled from PubMed searches, from the linked ORCID record, or by **adding a citation manually** (choose the closest non-journal type for the two software products). Up to five per list; drag to reorder. Worksheet section 4 gives the proposed ten plus verified alternates.
6. Click **View Draft** to check the layout, then **Download PDF**. SciENcv shows the certification statement; Brian clicks **CERTIFY** (a delegate can draft but cannot certify). Save the PDF into this folder and upload it in Research.gov under the proposal's Senior/Key Personnel → Biographical Sketch.

## ORCID warning

A previous session found that Brian's ORCID record (0000-0002-2748-1444) is missing works from 2022 onward. An ORCID import into SciENcv will therefore stop at 2021. Products published after that (in the worksheet: A2 DicomRTTool software, A5 elective-nodes paper, B3 EPID paper, B5 Keras contribution, and bench items C6, C7, C8) must either be **added to ORCID first** (ORCID → Works → Add → search by DOI, then re-sync in SciENcv) or **typed into My Bibliography manually** in SciENcv. Either way, check every imported citation against the DOI column of the worksheet before certifying.

## Before certifying

- Resolve every yellow `[CONFIRM]` in the worksheet (section 6 lists them).
- Use the same position-title wording here, in Current and Pending (Other) Support, in the Collaborators and Other Affiliations spreadsheet, and in Synergistic Activities.
- Make sure nothing personal (home address, phone) is in the document; extra content can get the proposal returned without review.

## Regenerating the worksheet PDF

Word COM export (same pattern as `..\..\.claude\skills\update-cv\`):

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("C:\Users\Markb\Modular_Projects\CV_and_Biosketch\Clincal_CV\NSF\1_Biosketch\NSF_Biosketch_SciENcv_Worksheet.docx")
$doc.SaveAs([ref]"C:\Users\Markb\Modular_Projects\CV_and_Biosketch\Clincal_CV\NSF\1_Biosketch\NSF_Biosketch_SciENcv_Worksheet.pdf", [ref]17)
$doc.Close($false)
$word.Quit()
```
