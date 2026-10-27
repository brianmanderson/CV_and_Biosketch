# NSF Synergistic Activities — Brian Mark Anderson

**Status: draft for Brian's review.** Nothing here is committed; review, edit the
`ITEMS` text in `build_synergistic_activities.py` (or the .docx directly), re-export
the PDF, and commit the .docx and .pdf together.

## Files

| File | Role |
|---|---|
| `Anderson_BM_NSF_Synergistic_Activities.docx` | Editable source, built by the script below |
| `Anderson_BM_NSF_Synergistic_Activities.pdf` | Word export of the .docx; the file that gets uploaded. Verified exactly one page with pypdf. |
| `build_synergistic_activities.py` | Rebuilds the .docx with python-docx. All wording lives in the `HEADER` and `ITEMS` constants. |
| `export_pdf.ps1` | Re-exports the PDF with Word COM and fails if it is not one page |

Regenerate after any edit:

```
python build_synergistic_activities.py .
powershell -ExecutionPolicy Bypass -File export_pdf.ps1
```

## Rules this document follows (PAPPG 24-1, Chapter II.D.2.h(iv))

- Synergistic Activities is a **separate one-page upload** for every senior/key person
  (since the May 2024 PAPPG, it is no longer part of the SciENcv biosketch). Upload the PDF.
- **One page maximum.**
- **Up to five distinct examples** of the broader impact of professional and scholarly
  activities: creation, integration, and transfer of knowledge. NSF's own example categories:
  innovations in teaching and training, contributions to the science of learning, development
  or refinement of research tools, computational methods and algorithms, databases supporting
  research and education, broadening participation of underrepresented groups, and service to
  the scientific community outside one's own organization.
- **Each example must be one specific activity.** NSF states that examples with multiple
  components (for instance a list of presentations, or every journal reviewed for) will not
  be considered. The draft therefore frames each item as a single activity plus the evidence
  of its impact. When tailoring, resist turning an item back into a list.
- **Format** (PAPPG 24-1 II.C.2): Times New Roman 11 pt or larger, 1-inch margins on all
  sides, single spacing, no more than six lines per vertical inch. The draft uses
  Times New Roman 11 pt, 1-inch margins, single spacing, US Letter.
- If a tailored version runs long, **tighten the prose**. Never shrink the margins or the font.

Current fill: 46 lines; the last line starts at 671 pt of the 720 pt text area, so roughly
three lines of Times New Roman 11 pt are spare.

## The five items and where every number comes from

All claims trace to `Clincal_CV/BMAnderson CV.docx` (section → entry).

| # | Item | Claims | CV source |
|---|---|---|---|
| 1 | Open-source software | DicomRTTool: Practical Radiation Oncology 2021; ~1,700 downloads/month; 100+ GitHub stars; 30+ forks; international academic and industry use | Major Contributions → "Open-Source Software: DicomRTTool"; Papers → Practical Radiation Oncology 02/2021 |
| 1 | | `reconstruct_patches` operation and ReconstructPatches2D/3D layers merged into Keras core (2026); backend-agnostic (TensorFlow, PyTorch, JAX); ~18 million monthly installs | Major Contributions → "Core Contributions to Keras/TensorFlow" |
| 1 | | DICOM Attribute Manipulation Tool, Journal of Applied Clinical Medical Physics 2025 | Papers → JACMP 04/2025; Clinical Projects (UNC) → "DICOM Manipulation Tools" |
| 2 | Nomenclature standards | AAPM TG-263 member since 2021 | Professional Service → AAPM → "2021 – Present Task Group 263" (see discrepancy below) |
| 2 | | Lead author, Open RT Structures, Red Journal 2024 | Major Contributions → "Nomenclature Standardization"; Papers → IJROBP 06/2024 |
| 2 | | Co-author, TG-263U1 roadblocks survey, JACMP 2024 | Papers → JACMP 06/2024 (Covington et al.) |
| 2 | | Adoption tool presented at AAMD Annual Conference 2023 | Oral Presentations → AAMD 06/2023 |
| 2 | | Technical Service Chief, CNS (2025–present); authored departmental CNS Standards and Guidelines | Professional Service → UCSD → "2025 – Present Technical Service Chief, Central Nervous System" |
| 3 | Translation to a trial | SIR Allied Scientist Grant, $40,000, 2019 | Grants and Fellowships → 2019 |
| 3 | | Biomechanical DIR for ablation outcomes, Medical Physics 2021 | Papers → Medical Physics 07/2021 |
| 3 | | COVER-ALL Phase 2 RCT, NCT04083378; software-based margin assessment superior to visual; Lancet Gastroenterology & Hepatology 2025 | Major Contributions → "Bench-to-Bedside"; Research Experience → PhD Dissertation; Papers → Lancet GH 05/2025 |
| 3 | | Alfred G. Knudson Jr. Outstanding Dissertation Award, 2022 | Honors and Awards → 2022 |
| 4 | Deep-learning training | Workshops at Winter Institute of Medical Physics (2020, 2025) and MD Anderson Image Guided Cancer Therapy Research Program (2020, 2021) | Invited Talks → 02/2020, 02/2025, 01/2020, 11/2021 |
| 4 | | AAPM Practical Big Data Workshop materials (GitHub/Jupyter guide; TensorFlow tfrecords pipeline), 2021; Early Career Investigator Impact Award | Poster Presentations → 09/2021 (two entries); Honors and Awards → 2021 |
| 4 | | Teaching Assistant, AAPM Summer School 2019 (workbooks) | Professional Service → AAPM → Other Service → 2019 |
| 4 | | Guest lectures: Rice University ELEC/COMP 576 (2019); Georgia Institute of Technology (2023) | Invited Talks → 09/2019, 11/2023 |
| 4 | | Chief Resident, UC San Diego, 2021–2023 | Clinical Projects → UCSD → "2021 – 2023 Chief Resident" |
| 5 | Service and broadening participation | Associate Editor, Medical Physics, and Board of Associate Editors (2023–present); ad hoc AE 2020–2021; reviewer since 2019 | Professional Service → Medical Physics (Journal); AAPM → MPBAE |
| 5 | | AAPM Machine Intelligence Subcommittee (2024–present); QUANTEC 2 Working Group (2026–present) | Professional Service → AAPM → Committees & Working Groups |
| 5 | | Medical Physicists for World Benefit member (2017–present) | Professional Service → MPWB |
| 5 | | Founding member and domestic liaison, International Students Association (2019–2021) | Professional Service → MD Anderson → ISA |
| 5 | | AAPM WizKids (2018–2019); UTHealth Science Night (2017–2019) | Professional Service → AAPM → Other Service; MD Anderson → Science Night |
| 5 | | Co-Investigator, AHRQ R01 (2024), AI-assisted sim review in rural settings | Grants and Fellowships → 2024 |

### Deliberately left out

- **"300+ liver segmentation cases deployed clinically by graduation"** was suggested for
  item 3 but appears nowhere in the CV, so it was not used. Add it only if Brian can source it
  (and consider adding it to the CV first).
- No DOIs: the CV does not carry them, and NSF does not require them. NCT04083378 is included
  because it is in the CV and strengthens item 3.

### Discrepancy to resolve in the CV

- TG-263 membership is dated **2021–2025** under Major Contributions but **2021 – Present**
  under Professional Service Activities. The draft says "since 2021" to be true to both.
  Pick one and fix the CV (and `Professional Service Activities.docx`, which mirrors that section).

## What to tailor per proposal

- **Header line:** optionally append the proposal title or NSF program. Not required.
- **Item 1, last sentence** (DICOM Attribute Manipulation Tool): first thing to cut if space is
  needed or the proposal is not imaging/DICOM focused.
- **Item 2:** for data/AI proposals lead with Open RT Structures; for clinical proposals lead with
  the CNS Standards and Guidelines. Both are already present; reorder rather than add.
- **Item 3:** for a software-to-clinic translation proposal, the trial protocol
  (CardioVascular and Interventional Radiology 2022) and the Radiology 2023 ablative-margin paper
  can be cited in place of Medical Physics 2021.
- **Item 4:** pick the guest lectures and workshop venues closest to the proposal's audience.
  Invited talks (ESTRO-AAPM 2023, AAPM Therapy Symposium 2024, MGH 2025, AAPM Symposium 2026)
  can replace a venue, but keep the item framed as one training activity, not a talk list.
- **Item 5:** the AHRQ R01 sentence overlaps with Current and Pending Support; keep it when the
  proposal touches rural or low-resource settings, otherwise replace it with the 2018 Medical
  Physics Slam (AAPM People's Choice; SWAAPM 1st place) as lay-audience science communication.
- **Numbers to refresh before each submission:** DicomRTTool downloads, stars, and forks;
  Keras monthly installs; every "present" date. Update the CV first so the two never disagree.
