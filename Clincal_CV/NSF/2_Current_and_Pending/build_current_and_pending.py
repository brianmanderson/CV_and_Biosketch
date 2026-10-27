"""Build NSF_Current_and_Pending_Worksheet.docx (Times New Roman 11 pt).

Content worksheet for Brian Mark Anderson's NSF Current and Pending (Other)
Support document. The official form must be generated and certified in SciENcv;
this worksheet holds the values to paste, one block per SciENcv entry, with the
public source for every verified value and [CONFIRM] on everything else.

Run from anywhere:  python -X utf8 build_current_and_pending.py
Output lands next to this script.
"""
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

HERE = Path(__file__).resolve().parent
OUT = HERE / "NSF_Current_and_Pending_Worksheet.docx"

FONT = "Times New Roman"
PREPARED = "10/27/2026"

# ---------------------------------------------------------------- content ---

OBJECTIVES_AHRQ = (
    "Develop, integrate, and evaluate an artificial intelligence/machine learning "
    "(AI/ML) intervention for pre-treatment peer review in radiation therapy (RT). "
    "Aim 1: develop AI/ML algorithms and visual summaries that characterize provider "
    "variability in key treatment-planning steps (e.g., target-volume definition and "
    "prescribed dose) and flag potential errors for reviewers, refined through "
    "user-centered design cycles. Aim 2: implement the integrated intervention in the "
    "clinical environment and evaluate its impact on patient safety with a "
    "cluster-randomized stepped-wedge design (ClinicalTrials.gov NCT07463833), "
    "quantifying safety as the rate per 1,000 RT fractions of clinically relevant "
    "errors missed at pre-treatment peer review and caught by downstream QA. "
    "Dr. Anderson's role: [CONFIRM - e.g., co-investigator leading development of "
    "the AI strategies and software for plan-quality and safety review, including "
    "tools suited to low-resource and rural settings]."
)
assert len(OBJECTIVES_AHRQ) <= 1500, len(OBJECTIVES_AHRQ)

SRC_REPORTER_24 = "https://reporter.nih.gov/project-details/10839671"
SRC_REPORTER_25 = "https://reporter.nih.gov/project-details/11062383"
SRC_AHRQ = ("https://digital.ahrq.gov/ahrq-funded-projects/"
            "development-and-assessment-artificial-intelligence-ai-enhanced-pretreatment")
SRC_UNC = ("https://www.med.unc.edu/intranet/2024/04/"
           "mazur-and-marks-receive-1-3-million-to-develop-ai-algorithms-to-enhance-patient-safety/")
SRC_PA = "https://grants.nih.gov/grants/guide/pa-files/PA-21-267.html"
SRC_PA24 = "https://grants.nih.gov/grants/guide/pa-files/PA-24-261.html"
SRC_CT = "https://clinicaltrials.gov/study/NCT07463833"
SRC_PAPPG = "https://www.nsf.gov/policies/pappg/24-1/ch-2-proposal-preparation#ch2D2hii"
SRC_FAQ = "https://www.nsf.gov/policies/document/faq-using-sciencv"
SRC_DISCL = "https://www.nsf.gov/policies/nspm-33/disclosures"
SRC_NLM = "https://www.nlm.nih.gov/ncbi/workshops/2024-04_SciENcv/NSF_Current_Pending_Other_Support.html"
SRC_HSRPROJ = "https://www.nlm.nih.gov/pubs/techbull/mj21/mj21_nichsr.html"

# ------------------------------------------------------------- helpers ------


def set_run_font(run, size=11, bold=None, italic=None, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)
    run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def style_font(style, size, bold=False, color=(0, 0, 0)):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(*color)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), FONT)
    # drop theme-font attributes that would override rFonts
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]


def para(doc, text="", bold=False, italic=False, size=11, style=None,
         space_after=6, align=None, keep_with_next=False):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    if text:
        r = p.add_run(text)
        set_run_font(r, size=size, bold=bold, italic=italic)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    if align:
        p.alignment = align
    if keep_with_next:
        p.paragraph_format.keep_with_next = True
    return p


def rich(doc, parts, style=None, space_after=6):
    """parts: list of (text, dict(bold=, italic=)) tuples in one paragraph."""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    for text, fmt in parts:
        r = p.add_run(text)
        set_run_font(r, size=11, **fmt)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    return p


def bullet(doc, text, parts=None):
    p = doc.add_paragraph(style="List Bullet")
    if parts:
        for t, fmt in parts:
            r = p.add_run(t)
            set_run_font(r, size=11, **fmt)
    else:
        r = p.add_run(text)
        set_run_font(r, size=11)
    p.paragraph_format.space_after = Pt(3)
    return p


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        set_run_font(r, size={1: 14, 2: 12, 3: 11}[level], bold=True, color=(0, 0, 0))
    h.paragraph_format.space_before = Pt({1: 14, 2: 10, 3: 8}[level])
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    return h


def shade(cell, hex_fill):
    tcpr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def cell_text(cell, text, bold=False, italic=False, size=11):
    """Write text into a cell; '\n' splits paragraphs. Marks [CONFIRM] bold."""
    cell.text = ""
    first = True
    for line in text.split("\n"):
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        # bold any [CONFIRM...] token so it stands out
        idx = 0
        while True:
            s = line.find("[CONFIRM", idx)
            if s < 0:
                seg = line[idx:]
                if seg:
                    set_run_font(p.add_run(seg), size=size, bold=bold, italic=italic)
                break
            if s > idx:
                set_run_font(p.add_run(line[idx:s]), size=size, bold=bold, italic=italic)
            e = line.find("]", s)
            e = len(line) if e < 0 else e + 1
            set_run_font(p.add_run(line[s:e]), size=size, bold=True, italic=italic)
            idx = e


def set_col_widths(table, widths):
    table.autofit = False
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w
    for i, w in enumerate(widths):
        table.columns[i].width = w


def repeat_header(row):
    trpr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trpr.append(el)


def field_table(doc, rows, header=("SciENcv field", "Value to enter", "Source / status")):
    """rows: list of (field, value, source)."""
    t = doc.add_table(rows=1, cols=3)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = (Inches(1.45), Inches(3.15), Inches(1.9))
    hdr = t.rows[0]
    repeat_header(hdr)
    for i, h in enumerate(header):
        cell_text(hdr.cells[i], h, bold=True)
        shade(hdr.cells[i], "D9D9D9")
    for field, value, source in rows:
        r = t.add_row()
        cell_text(r.cells[0], field, bold=True)
        cell_text(r.cells[1], value)
        cell_text(r.cells[2], source)
    set_col_widths(t, widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def simple_table(doc, header, rows, widths):
    t = doc.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = t.rows[0]
    repeat_header(hdr)
    for i, h in enumerate(header):
        cell_text(hdr.cells[i], h, bold=True)
        shade(hdr.cells[i], "D9D9D9")
    for row in rows:
        r = t.add_row()
        for i, v in enumerate(row):
            cell_text(r.cells[i], v, bold=(i == 0))
    set_col_widths(t, widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def add_page_number_footer(section, left_text):
    footer = section.footer
    p = footer.paragraphs[0]
    p.text = ""
    r = p.add_run(left_text + "    Page ")
    set_run_font(r, size=9, italic=True)
    # PAGE field
    r2 = p.add_run()
    set_run_font(r2, size=9, italic=True)
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if tag:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), tag)
            r2._element.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            r2._element.append(it)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ---------------------------------------------------------------- build -----


def build():
    doc = Document()

    # page setup: Letter, 1" margins
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Inches(1))

    # fonts
    style_font(doc.styles["Normal"], 11)
    doc.styles["Normal"].paragraph_format.space_after = Pt(6)
    for name, size in (("Heading 1", 14), ("Heading 2", 12), ("Heading 3", 11)):
        style_font(doc.styles[name], size, bold=True)
    style_font(doc.styles["List Bullet"], 11)
    style_font(doc.styles["Title"], 16, bold=True)

    doc.core_properties.title = "NSF Current and Pending (Other) Support - Content Worksheet"
    doc.core_properties.author = "Brian Mark Anderson"
    doc.core_properties.subject = "Draft content for SciENcv; not an NSF form"

    add_page_number_footer(
        sec, "DRAFT worksheet for review - not an NSF document. The official form is generated and certified in SciENcv.")

    # ---------------------------------------------------------------- title
    para(doc, "NSF Current and Pending (Other) Support", bold=True, size=16, space_after=0)
    para(doc, "Content worksheet for SciENcv", bold=True, size=13, space_after=2)
    para(doc, "Brian Mark Anderson, PhD, DABR  |  Clinical Assistant Professor of Medical Physics, "
              "University of California, San Diego  |  eRA Commons: BrianMAnderson  |  "
              "ORCID 0000-0002-2748-1444", size=10, space_after=2)
    para(doc, f"Prepared {PREPARED}. DRAFT for Brian's review: nothing here is final until he confirms it.",
         italic=True, size=10, space_after=8)

    heading(doc, "How to use this worksheet", 1)
    rich(doc, [
        ("NSF requires every senior/key person's Current and Pending (Other) Support to be prepared in SciENcv, "
         "which produces the only NSF-compliant PDF (PAPPG 24-1, Chapter II.D.2.h(ii)). ", {}),
        ("This worksheet is the content, not the form. ", {"bold": True}),
        ("Each block below mirrors one SciENcv entry, field by field, so the values can be pasted in order. "
         "The right-hand column of every table says where a value came from; ", {}),
        ("[CONFIRM]", {"bold": True}),
        (" marks anything that could not be verified from a public source and must come from Brian or a "
         "sponsored-projects office before it is entered. The companion README.md lists the exact questions "
         "to ask the UNC and UCSD offices and where each answer goes.", {}),
    ])
    rich(doc, [
        ("Rules that shape every entry (PAPPG 24-1 II.D.2.h(ii); NSF SciENcv FAQ):", {"bold": True}),
    ], space_after=2)
    bullet(doc, "", [("Current", {"bold": True}),
                     (" = any active project or project with ongoing obligations, from any source (federal, state, "
                      "foreign, foundation, industry, or internal institutional funds). ", {}),
                     ("Pending", {"bold": True}),
                     (" = any proposal under consideration, including this NSF proposal itself.", {})])
    bullet(doc, "", [("Only projects Brian commits time to are listed. ", {}),
                     ("SciENcv rejects a person-months value of 0", {"bold": True}),
                     (" (FAQ Q23): if his effort on a project is now zero, the project is removed, not zero-filled. "
                      "Enter the real figure, including unsalaried effort; estimates are acceptable (minimum 0.01).", {})])
    bullet(doc, "", [("Total award amount", {"bold": True}),
                     (" is the whole-dollar total for the full award period including indirect costs. "
                      "Dates are MM/YYYY. ", {}),
                     ("Overall objectives", {"bold": True}),
                     (" is limited to 1,500 characters. ", {}),
                     ("Overlap statement", {"bold": True}),
                     (": write \"None\" when there is no overlap in scope, budget, or person-months.", {})])
    bullet(doc, "", [("No personal information", {"bold": True}),
                     (" (home address, personal phone/e-mail) goes in the form.", {})])
    bullet(doc, "", [("Certification cannot be delegated:", {"bold": True}),
                     (" Brian ticks the certification in his own SciENcv account before the PDF can be downloaded. "
                      "The downloaded PDF must be submitted as-is (renaming is fine; opening and re-saving in another "
                      "application is not).", {})])

    # -------------------------------------------------------- identifying info
    heading(doc, "SciENcv Section 1 - Identifying information, organization and location", 1)
    field_table(doc, [
        ("Name", "Brian Mark Anderson", "CV header."),
        ("Position title", "Clinical Assistant Professor of Medical Physics",
         "CV header. The NIH biosketch says \"Department of Radiation Oncology\"; the CV says \"Radiation Medicine "
         "& Applied Sciences\". Use the department name on the UCSD appointment letter [CONFIRM]."),
        ("Organization", "University of California, San Diego - Department of Radiation Medicine and Applied Sciences",
         "CV. [CONFIRM department wording with UCSD]"),
        ("City / State / Country", "La Jolla, CA, USA",
         "The UCSD campus mailing address is La Jolla (9500 Gilman Drive); the CV header says San Diego. Match "
         "whatever the proposal cover sheet uses for the performance site [CONFIRM]."),
        ("Sign-in identifiers (not entered in the form)", "eRA Commons BrianMAnderson; ORCID 0000-0002-2748-1444",
         "CLAUDE.md. Either route lands in the same NCBI account once linked."),
    ])

    # --------------------------------------------------------- entry A (AHRQ)
    heading(doc, "SciENcv Section 2 - Projects and proposals", 1)
    heading(doc, "Entry A - CURRENT: AHRQ R18 HS029474 (UNC Chapel Hill; Mazur / Marks)", 2)
    rich(doc, [
        ("What the public record shows. ", {"bold": True}),
        ("The CV and NIH biosketch list this as a 2024 \"R-01, Co-Investigator\" under \"AHRQ Health Services "
         "Research Projects: Examining the Impact of Artificial Intelligence on Healthcare Safety\". NIH RePORTER "
         "and AHRQ's own project listing show the Mazur/Marks award is an ", {}),
        ("R18", {"bold": True}),
        (" (Research Demonstration and Dissemination Project), number ", {}),
        ("R18 HS029474", {"bold": True}),
        (", awarded 03/28/2024 under PA-21-267 \"Making Health Care Safer in Ambulatory Care Settings and "
         "Long-term Care Facilities (R18)\". \"Examining the Impact of Artificial Intelligence (AI) on Healthcare "
         "Safety\" is the title of a different AHRQ announcement (PA-24-261, also R18). RePORTER lists only the "
         "PIs, so Brian's Co-Investigator role and effort cannot be verified publicly; they must come from the UNC "
         "Office of Sponsored Programs.", {}),
    ])
    field_table(doc, [
        ("Status", "Current",
         "[CONFIRM] Keep only if Brian still commits person-months after his 07/2025 move to UCSD (UNC adjunct "
         "appointment 2025-present). If his effort is now 0, omit the entry (FAQ Q23)."),
        ("Title",
         "Development and Assessment of Artificial Intelligence (AI)-Enhanced Pre-treatment Peer-review Process "
         "to Improve Patient Safety in Radiation Oncology",
         "NIH RePORTER project title (appl. 10839671). AHRQ's listing spells \"Pretreatment\" as one word; "
         "either is acceptable."),
        ("Award number",
         "R18HS029474\n(original award 1R18HS029474-01A1; current budget year 5R18HS029474-02)",
         "NIH RePORTER. Not an R01: the CV and NIH biosketch need correcting (separate follow-up). "
         "[CONFIRM with UNC OSP that no separate R01 lists Brian as Co-I]"),
        ("Source of support",
         "Agency for Healthcare Research and Quality (AHRQ), U.S. Department of Health and Human Services. "
         "Prime awardee: University of North Carolina at Chapel Hill (contact PI Lukasz Mazur, PhD; MPI "
         "Lawrence B. Marks, MD).",
         "NIH RePORTER; AHRQ project page. [CONFIRM] If Brian's post-2025 effort runs through a UNC-to-UCSD "
         "subaward, write \"AHRQ (prime: University of North Carolina at Chapel Hill; subaward to UC San Diego)\"."),
        ("Primary place of performance", "Chapel Hill, NC, USA",
         "RePORTER organization record. [CONFIRM] If Brian's work is performed at UCSD under a subaward, enter "
         "La Jolla, CA, USA instead."),
        ("Start date", "04/2024", "RePORTER and AHRQ: project period 04/01/2024 - 03/31/2027."),
        ("End date", "03/2027",
         "Same. [CONFIRM] whether a no-cost extension is planned; if so use the extended end date."),
        ("Total anticipated amount (incl. indirect costs)", "$1,345,650",
         "AHRQ project page total. RePORTER by budget year: Y1 (04/2024-03/2025) $450,000 = $307,898 direct + "
         "$142,102 indirect; Y2 (04/2025-03/2026) $450,000 = $312,529 + $137,471; Y3 not yet posted. UNC's "
         "announcement rounds to \"$1.3 million\". [CONFIRM] If a subaward exists, the sponsored-projects office "
         "may instruct entering the subaward total to UCSD instead; follow their instruction."),
        ("Person-months per budget year (current period first)",
         "Budget year 3 (04/2026 - 03/2027): [CONFIRM x.xx calendar months]\n"
         "Add a row for 04/2027 - 03/2028 only if a no-cost extension is granted.\n"
         "Prior years (Y1 04/2024-03/2025, Y2 04/2025-03/2026) are not entered; SciENcv asks for the current and "
         "each subsequent period.",
         "From UNC OSP / the award budget. Enter calendar months, including unsalaried effort. SciENcv rejects 0 "
         "(FAQ Q23); minimum 0.01. For an April-March budget year SciENcv's \"Year\" field takes the later "
         "calendar year (FAQ Q24), i.e. 2027 for the current period."),
        ("Overall objectives (max 1,500 characters)",
         OBJECTIVES_AHRQ,
         f"Drafted from the RePORTER project abstract and the AHRQ project summary; {len(OBJECTIVES_AHRQ)} "
         "characters including the role placeholder. Replace the bracketed role sentence with Brian's actual "
         "responsibilities as stated in the award [CONFIRM]."),
        ("Statement of potential overlap", "None",
         "[CONFIRM] once the NSF aims are final. If the NSF proposal also develops AI tools for radiotherapy "
         "plan-quality or safety review, replace with one sentence that distinguishes the two scopes and states "
         "that the budgets and person-months do not overlap."),
    ])
    rich(doc, [("Public sources for Entry A", {"bold": True})], space_after=2)
    bullet(doc, f"NIH RePORTER, FY2024 record 1R18HS029474-01A1 (title, dates, PIs, abstract, costs): {SRC_REPORTER_24}")
    bullet(doc, f"NIH RePORTER, FY2025 record 5R18HS029474-02 (budget period 04/2025-03/2026): {SRC_REPORTER_25}")
    bullet(doc, f"AHRQ Digital Healthcare Research project listing (total $1,345,650; period 04/01/2024-03/31/2027): {SRC_AHRQ}")
    bullet(doc, f"UNC School of Medicine announcement, 04/2024 (\"$1.3 million\", three-year project): {SRC_UNC}")
    bullet(doc, f"Funding opportunity the award was made under, PA-21-267 (R18; caps of $450,000/year and $1.35 M total, 3 years): {SRC_PA}")
    bullet(doc, f"The announcement whose title the CV uses, PA-24-261 \"Examining the Impact of AI on Healthcare Safety (R18)\" (a different NOFO): {SRC_PA24}")
    bullet(doc, f"ClinicalTrials.gov NCT07463833, the Aim 2 stepped-wedge trial (sponsor UNC Lineberger; started 06/2026): {SRC_CT}")
    bullet(doc, f"HSRProj could not be searched: NLM retired it on 09/14/2021: {SRC_HSRPROJ}")
    bullet(doc, "Not findable publicly: Brian's named role, his budgeted effort, and any subaward. RePORTER lists PIs only; "
                "the AHRQ page and ClinicalTrials.gov name no other team members.")

    rich(doc, [("Possible additional Pending entries from the same group [CONFIRM]", {"bold": True})], space_after=2)
    rich(doc, [
        ("The CV entry also lists \"Systems-Based Approaches to Improve Patient Safety\" and \"Efficacy of AI assisted "
         "sim review in rural settings: randomized control study\". The first reads like an AHRQ announcement title "
         "rather than a project; the second like a distinct study. If any proposal from the Mazur/Marks group that "
         "names Brian as senior/key personnel is still under review (for example a submission to PA-24-261), it must "
         "be entered as a separate ", {}),
        ("Pending", {"bold": True}),
        (" block with the same fields: title, proposal number if assigned, source, place, requested dates, requested "
         "total including indirects, person-months per year, objectives, and overlap. Ask UNC OSP (README, question 6).", {}),
    ])

    # ------------------------------------------------------ entry B (NSF pending)
    heading(doc, "Entry B - PENDING: this NSF proposal", 2)
    rich(doc, [
        ("PAPPG 24-1 defines Pending as any proposal under consideration, ", {}),
        ("including the proposal being submitted", {"bold": True}),
        (". Every value below comes from the proposal's cover sheet and budget; none is in the CV.", {}),
    ])
    field_table(doc, [
        ("Status", "Pending", "Required for the proposal itself."),
        ("Title", "[CONFIRM - exact proposal title as entered on the Research.gov cover sheet]", "Proposal cover sheet."),
        ("Proposal number", "Leave blank (not yet assigned)",
         "NSF assigns the number at submission. If the proposal has already been submitted and NSF has issued a "
         "number, enter it [CONFIRM]."),
        ("Source of support", "National Science Foundation [CONFIRM directorate / division / program solicitation]",
         "Cover sheet."),
        ("Primary place of performance", "La Jolla, CA, USA [CONFIRM]",
         "Cover sheet \"Primary Place of Performance\". Use the same city as Section 1."),
        ("Start date", "[CONFIRM MM/YYYY - requested start]", "Cover sheet."),
        ("End date", "[CONFIRM MM/YYYY - requested end]", "Cover sheet."),
        ("Total anticipated amount (incl. indirect costs)", "$[CONFIRM - total requested, line L of the budget]",
         "Budget justification / Research.gov budget. Whole dollars, including F&A for the full requested period."),
        ("Person-months per budget year",
         "Year 1 ([CONFIRM MM/YYYY-MM/YYYY]): [x.xx]\nYear 2: [x.xx]\nYear 3: [x.xx]\n(add or remove years to match the requested period)",
         "Budget pages (calendar, academic, or summer months as budgeted). Include committed unsalaried effort. "
         "SciENcv \"Year\" = later calendar year of each budget period (FAQ Q24)."),
        ("Overall objectives (max 1,500 characters)",
         "[CONFIRM - paste the proposal's overview/objectives, trimmed to 1,500 characters]",
         "Project Summary (Overview) is the usual source."),
        ("Statement of potential overlap", "None",
         "If the NSF work builds on the AHRQ project, say so explicitly and state that the scope, budget, and "
         "person-months do not overlap [CONFIRM]."),
    ])

    # ------------------------------------------------------- other candidates
    heading(doc, "Other sources of support to rule in or out [CONFIRM]", 2)
    rich(doc, [("Each item below is reportable only if it is active or pending, supports Brian's research, and "
                "carries a time commitment. The CV suggests none applies, but only Brian can confirm.", {})])
    simple_table(doc, ("Candidate", "What the rule says", "Worksheet action"), [
        ("UCSD internal funds (departmental seed or pilot grants, internal awards)",
         "Internal institutional funding counts as a source of support when it funds research with an effort commitment.",
         "If any exist, add a Current block with the same fields; source = \"University of California, San Diego (internal)\"."),
        ("UCSD startup package",
         "Startup packages from the proposing organization are NOT reported (NSF disclosure table).",
         "Nothing to enter."),
        ("UNC startup package (2023-2025 appointment)",
         "Startup packages from an organization other than the proposer ARE reported if still active.",
         "Not current after the 2025 departure unless funds remain accessible [CONFIRM]; if so, add a Current block."),
        ("Industry research agreements, vendor-sponsored projects (Varian, Elekta, RaySearch, etc.)",
         "Reportable when they fund research and involve Brian's time; equipment or data loans fall under in-kind below.",
         "None in the CV. Add a block if any exist."),
        ("Other pending proposals (any sponsor, including NIH/AHRQ or foundation submissions)",
         "Every proposal under consideration on which Brian is senior/key personnel is Pending.",
         "Add a Pending block per proposal."),
    ], (Inches(1.9), Inches(2.6), Inches(2.0)))

    # -------------------------------------------------------- in-kind section
    heading(doc, "SciENcv Section 3 - In-kind contributions", 1)
    rich(doc, [
        ("Rule (PAPPG 24-1 II.D.2.h(ii)). ", {"bold": True}),
        ("Report non-cash contributions from an entity outside UCSD that directly support Brian's research, require a "
         "commitment of his time, and have an estimated value of ", {}),
        ("$5,000 or more", {"bold": True}),
        (" (examples: laboratory or office space, equipment, data or datasets, supplies, goods and services, employee "
         "or student resources). Contributions under $5,000, or with no associated time commitment, are not reported. "
         "In-kind resources intended for use on this NSF project go in the proposal's Facilities, Equipment and Other "
         "Resources section instead, not here.", {}),
    ])
    field_table(doc, [
        ("Status", "Current or Pending", "Per contribution."),
        ("Receipt date (or anticipated)", "MM/YYYY", ""),
        ("Source", "Name of the entity providing the contribution", ""),
        ("Summary of the contribution", "What was provided", ""),
        ("Person-months per year", "Calendar months of Brian's time the contribution requires, current period first", "Minimum 0.01."),
        ("U.S. dollar value", "Estimated value, whole dollars", "Must be $5,000 or more to be reportable."),
        ("Overall objectives (max 1,500 characters)", "What the contribution supports", ""),
        ("Statement of potential overlap", "\"None\" if none", ""),
    ], header=("SciENcv field", "What goes there", "Note"))
    rich(doc, [
        ("Determination for this worksheet: NONE IDENTIFIED. ", {"bold": True}),
        ("Nothing in the CV describes an external contribution of space, equipment, data, or personnel that requires "
         "Brian's time. Items considered and ruled out: UCSD departmental resources (from the proposing organization, "
         "not reported); the UNC adjunct appointment (an appointment, not a resource - see the next section); AAPM "
         "committee and journal roles (unpaid service, no resources); open-source software collaborations such as "
         "DicomRTTool and the Keras contribution (no external resources received). ", {}),
        ("[CONFIRM]", {"bold": True}),
        (" that no collaborator supplies datasets, compute, equipment, or staff for Brian's research at or above "
         "$5,000 - the UNC AHRQ team's data access is part of the funded award, not a separate in-kind contribution, "
         "unless UNC provides it outside the award.", {}),
    ])

    # ------------------------------------------- appointments vs support
    heading(doc, "Appointments are not support (biosketch items, not entries here)", 1)
    rich(doc, [
        ("The NSF disclosure table places academic, professional, or institutional appointments - paid or unpaid - in "
         "the ", {}),
        ("Biographical Sketch", {"bold": True}),
        (", not in Current and Pending Support. Current and Pending Support lists resources (funding, in-kind, "
         "consulting), so the following roles from the CV stay out of this document ", {}),
        ("unless they carry salary or research resources from a source outside UCSD", {"bold": True}),
        (", in which case that funding becomes an entry:", {}),
    ])
    bullet(doc, "UCSD: Chair of Technical Innovation and Implementation (2025 - present); Technical Service Chief, "
                "Central Nervous System (2025 - present); Search Committee Member (2026 - 2027). Departmental roles "
                "within the proposing organization: biosketch only.")
    bullet(doc, "UNC Chapel Hill: Adjunct Assistant Professor (2025 - present). Unpaid adjunct appointment: biosketch "
                "only. [CONFIRM] it carries no salary; if UNC pays Brian for AHRQ effort through this appointment, "
                "that is the Entry A effort, not a separate entry.")
    bullet(doc, "AAPM: Task Group 263, Machine Intelligence Subcommittee, QUANTEC 2 Working Group, Medical Physics "
                "Board of Associate Editors; Medical Physics journal Associate Editor (2023 - present). Unpaid service: "
                "biosketch only. [CONFIRM] no honorarium tied to research effort.")

    # ------------------------------------------------------ consulting note
    heading(doc, "Consulting", 1)
    rich(doc, [
        ("Rule (PAPPG 24-1 II.D.2.h(ii); NSF disclosure table). ", {"bold": True}),
        ("Consulting is entered in the projects and active projects section when it (1) requires Brian to perform "
         "research, (2) does not involve research but is related to his research portfolio and could affect his "
         "funding, effort, or scientific integrity, or (3) is under a contract requiring him to conceal or withhold "
         "confidential financial ties, regardless of its duration. ", {}),
        ("The dollar amount Brian receives goes in the total-award-amount field", {"bold": True}),
        (", with the company as the source, the contract dates, person-months, a one-line objectives statement, and "
         "an overlap statement.", {}),
    ])
    rich(doc, [
        ("Determination for this worksheet: NONE IDENTIFIED in the CV. ", {"bold": True}),
        ("[CONFIRM]", {"bold": True}),
        (" that Brian holds no paid consulting, advisory-board, or expert-witness engagement (including with "
         "radiotherapy vendors or AI companies) that meets any of the three tests. Purely clinical or teaching "
         "consulting unrelated to his research portfolio is not reported.", {}),
    ])

    # ------------------------------------------------------- closed awards
    heading(doc, "Closed awards - not reportable", 1)
    rich(doc, [("\"Current\" means active or with ongoing obligations. The following CV items are closed and are "
                "not entered anywhere in the form:", {})])
    simple_table(doc, ("Award (CV)", "Period and amount", "Why not reported"), [
        ("Society of Interventional Radiology (SIR) Allied Scientist Grant",
         "2019 - 2021 (two years); $40,000 total",
         "Ended ~2021. [CONFIRM] no open final-report obligation."),
        ("MD Anderson Dr. John J. Kopchick Fellowship", "2020, 2021; $15,000 per year", "Closed with the PhD (2021)."),
        ("AAPM Summer Undergraduate Fellowship", "2014; 10-week program", "Closed."),
        ("Georgia Tech President's Undergraduate Research Award (PURA)", "2013, 2014", "Closed."),
    ], (Inches(2.3), Inches(2.0), Inches(2.2)))

    # ------------------------------------------------- what Brian must supply
    heading(doc, "What Brian must supply before anything is entered", 1)
    simple_table(doc, ("Item", "Who has it", "Goes in"), [
        ("Whether he is still Co-I on R18 HS029474 with non-zero effort after the 2025 move; whether a UNC-to-UCSD subaward exists",
         "UNC Office of Sponsored Programs (OSP) / the Mazur-Marks research administrator; UCSD HS SPPO if a subaward",
         "Entry A: Status, Source of support, Place of performance"),
        ("Person-months for budget year 3 (04/2026 - 03/2027) and any extension year",
         "UNC OSP / award budget", "Entry A: Person-months"),
        ("Confirmation of the total award ($1,345,650) or the subaward total to UCSD, and the budget-period dates",
         "UNC OSP", "Entry A: Total amount, Start/End"),
        ("Brian's stated responsibilities on the award", "Award budget justification / Dr. Mazur", "Entry A: Objectives (role sentence)"),
        ("Any pending proposals from the UNC group or elsewhere naming him as senior/key personnel",
         "UNC OSP; Brian", "Additional Pending blocks"),
        ("NSF proposal title, requested dates, total requested including F&A, person-months per year, program",
         "UCSD Office of Contract and Grant Administration (OCGA) / department fund manager; the proposal budget",
         "Entry B: all fields"),
        ("Department name and performance-site city as the cover sheet states them", "UCSD OCGA / fund manager",
         "Section 1; Entry B: Place of performance"),
        ("Any UCSD internal funding, industry agreements, consulting, or in-kind contributions", "Brian",
         "New blocks as applicable; Consulting; In-kind"),
        ("Overlap wording once the NSF aims are final", "Brian", "Entries A and B: Overlap statement"),
    ], (Inches(2.7), Inches(2.0), Inches(1.8)))

    # ----------------------------------------------------------- rule sources
    heading(doc, "Rule sources", 1)
    bullet(doc, f"PAPPG 24-1, Chapter II.D.2.h(ii) Current and Pending (Other) Support: {SRC_PAPPG}")
    bullet(doc, f"NSF FAQ on using SciENcv (delegates, certification, zero person-months, fiscal-year entry, PDF handling): {SRC_FAQ}")
    bullet(doc, f"NSF NSPM-33 required-disclosures table (appointments vs support, in-kind $5,000 threshold, consulting tests, startup packages): {SRC_DISCL}")
    bullet(doc, f"NLM SciENcv training page for the NSF Current and Pending form (section layout and click path): {SRC_NLM}")
    bullet(doc, "PAPPG version note: NSF 24-1 (effective 05/20/2024) is the current guide as of this worksheet; NSF lists two "
                "policy supplements (NSF 26-200, NSF 26-202) that apply to awards made on or after 12/08/2025 and 01/22/2026. "
                "Check the PAPPG page for a newer edition before submission.")

    doc.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
