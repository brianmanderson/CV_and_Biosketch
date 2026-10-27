# -*- coding: utf-8 -*-
"""Merge the seven affiliation-lookup batch tables with the CV roster, apply corrections,
and emit (a) table_data.json for the fill script and (b) COA_source_list.md for Brian."""
import glob, json, os, re, sys
from roster import ENTRIES, roster, canon

HERE = os.path.dirname(os.path.abspath(__file__))
LOOKUPS = os.path.join(HERE, "..", "lookups")
OUT_JSON = os.path.join(HERE, "..", "table_data.json")
OUT_MD = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "COA_source_list.md")

REF_DATE = "10/27/2026"

# ---------------------------------------------------------------- batch tables
lookup = {}
for path in sorted(glob.glob(os.path.join(LOOKUPS, "batch*.md"))):
    batch = os.path.basename(path)
    for line in open(path, encoding="utf-8"):
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) < 6 or cells[0] in ("Key", "") or set(cells[0]) <= set("-: "):
            continue
        key, full, org, url, conf, notes = cells[:6]
        lookup[key] = dict(full=full, org=org, url=url, conf=conf.lower().split()[0], notes=notes, batch=batch)

# ---------------------------------------------------------------- corrections
# Keys that are the same person as another key (CV variant/typo); merged, not listed twice.
DROP = {"Arcot R": "Rajasekar A"}

# Display-name overrides: key -> "Surname, Given" (used where auto-derivation is wrong or the CV key is a typo)
NAME_OVERRIDE = {
    "Cao HT": "Tran Cao, Hop S.",
    "Sousa C": "de Sousa, Cecilia F. P. M.",
    "Jones AK": "Jones, A. Kyle",
    "Shan KY": "Shah, Ketan Y.",
    "Moncio A": "Moncion, Alexander",
    "Krupein A": "Krupien, Andrew J.",
    "Maxin P": "Maxim, Peter G.",
    "Knausl B": "Knäusl, Barbara",
    "Chen S": "Chen, S.",
    "Friend D": "Friend, D.",
    "Dong F": "Dong, F.",
    "Van Dijk L": "van Dijk, Lisanne V.",
}

# Organization overrides for rows the lookups left UNKNOWN or ambiguous: key -> (org, optional-note)
ORG_OVERRIDE = {
    "Gupta A": ("The University of Texas MD Anderson Cancer Center", "last known (Ph.D. 2026, Imaging Physics); current employer not found"),
    "Awais M": ("The University of Texas MD Anderson Cancer Center", "last known (Imaging Physics, 2024); current employer not found"),
    "Chiang TYD": ("The University of Texas MD Anderson Cancer Center", "last known (Diagnostic Imaging, 2022); no newer record found"),
    "McCulloch MM": ("The University of Texas MD Anderson Cancer Center", "last known (Imaging Physics, 2022); current employer not found"),
    "Cazoulat G": ("The University of Texas MD Anderson Cancer Center", "last known (Imaging Physics, 2024); current employer not found"),
    "Gryshkevych S": ("RaySearch Laboratories AB", "last known (2022); no newer record found"),
    "Mathew RT": ("Swedish Medical Center", "per Koay-lab alumni page ('Clinical Research Coordinator II, Swedish'); unverified"),
    "Chen S": ("University of North Carolina at Chapel Hill", "identity not resolved; co-authors all UNC Radiation Oncology (AAPM 2024 posters)"),
    "Friend D": ("University of North Carolina at Chapel Hill", "identity not resolved; co-authors all UNC Radiation Oncology (AAPM 2026 abstract)"),
    "Dong F": ("University of North Carolina at Chapel Hill", "first name not found; ASTRO 2026 listing gives UNC Radiation Oncology"),
    "Rigaud B": ("The University of Texas MD Anderson Cancer Center", "2025-26 bylines say MD Anderson; one unverified listing says Promaton (Amsterdam)"),
    "Patel S": ("University of North Carolina at Chapel Hill", "probably Shetal Patel (UNC Medical Oncology); if Samip Patel, it is Mayo Clinic Florida"),
    "Hsu D": ("Memorial Sloan Kettering Cancer Center", "identity inferred (Dylan G. Hsu); not confirmed against the AAPM 2026 listing"),
    "Kim S": ("University of California, San Diego", "identity inferred (Sangwoo Shawn Kim, MD); could be a different S. Kim on staff"),
    "Moore K": ("Varian Medical Systems, Inc.", "2024-25 bylines list Varian (one also UC San Diego); absent from 2026 UCSD faculty directory"),
    "Sousa C": ("Johns Hopkins University", "ASTRO 2026 listing prints first author as 'C. D. S. Smith'; same person under a changed surname?"),
}

# Short, explicit reasons for the medium/low-confidence rows that are not in ORG_OVERRIDE
FLAG_REASON = {
    "Briones-Dimayuga M": "research staff; only evidence is the 2025 Lancet byline",
    "Estes C": "newest evidence 2024-25 bylines (Mercy Springfield); no 2026 source",
    "Fellman BM": "Baylor directory lists him; 2026 bylines still MD Anderson Biostatistics (recent move)",
    "Foster M": "newest evidence 2025 (UNC Division of Healthcare Engineering); no 2026 source",
    "Gay S": "Ph.D. student in Court lab through early 2026; may have started a residency elsewhere",
    "Gebre T": "likely student; last evidence 01/2026 IJROBP byline; not on the ASTRO 2026 listing",
    "Hagan E": "likely student; last evidence 01/2026 IJROBP byline; not on the ASTRO 2026 listing",
    "Obaideen A": "likely student; last evidence 01/2026 IJROBP byline; not on the ASTRO 2026 listing",
    "Negassa M": "likely student; last evidence 01/2026 IJROBP byline; not on the ASTRO 2026 listing",
    "Karunaker A": "trainee; evidence is the 05/2026 IEEE CAI byline; printed Anirudh/Aniruddh",
    "Katz MS": "no 2026 source; 2025 byline Lowell General; also Radiation Oncology Associates (Elliot Hospital, NH)",
    "Liu CC": "2026 byline UNC; 2025 bylines and ORCID say Georgetown University",
    "Kirimli E": "per Koay-lab alumni page (family medicine physician, Turkiye Ministry of Health); unverified",
    "Zaid M": "per Koay-lab alumni page (Global Clinical Lead, AstraZeneca); latest byline still MD Anderson",
    "McHugh D": "ORCID and JACMP 2025: Physician Liaison, Curi (Raleigh, NC); record last updated 2023",
    "Mayo C": "one 10/2026 byline adds Brigham and Women's/Dana-Farber; a move may be in progress",
    "Valdes G": "2026 bylines Moffitt; a 2026 speaker bio says 'previously Moffitt' (OncoBrain AI); may have moved",
    "Krupein A": "CV spelling 'Krupein' not found anywhere; Andrew J. Krupien (UCLA) is the likely match",
    "Maxin P": "CV spelling 'Maxin' not found anywhere; Peter G. Maxim (UC Irvine) is the likely match",
}

# Details that add nothing in the Optional column
DETAIL_DROP = {"Mercy", "Tufts Medicine", "Northwell", "New Delhi, India", "Stockholm, Sweden", "Raleigh, NC", "Palo Alto, CA",
               "per 2025–2026 paper bylines; see note", "Anschutz Medical Campus", "Prof."}
DETAIL_OVERRIDE = {
    "Adapa K": "Regional Adviser, Digital Health; also Adjunct Faculty, UNC Chapel Hill",
    "Moore K": "",
    "Hong DS": "Department of Radiation Oncology (radiation oncologist)",
    "Price J": "Department of Radiation Oncology",
    "Zegers B": "Department of Radiation Medicine and Applied Sciences",
    "Chera B": "Department of Radiation Oncology, Hollings Cancer Center",
    "Vaccarelli M": "",
    "Morse R": "Department of Radiation Oncology",
}
# (organization, detail) fixes applied after the generic split for high-confidence rows
ORG_FIX = {
    "Ritter T": ("Virginia Commonwealth University Health System", "Department of Radiation Oncology; also VA Richmond Healthcare System"),
    "Simonds H": ("University Hospitals Plymouth NHS Trust", "Plymouth, United Kingdom"),
    "Reber B": ("Mayo Clinic", "Rochester, MN; medical physics resident"),
}

# Normalize the big institutions to one spelling each
ORG_NORMALIZE = [
    (re.compile(r"^University of California,? San Diego.*$|^UC San Diego.*$", re.I), "University of California, San Diego"),
    (re.compile(r"^University of North Carolina at Chapel Hill.*$", re.I), "University of North Carolina at Chapel Hill"),
    (re.compile(r"^The University of Texas MD Anderson Cancer Center.*$", re.I), "The University of Texas MD Anderson Cancer Center"),
    (re.compile(r"^Johns Hopkins University.*$", re.I), "Johns Hopkins University"),
    (re.compile(r"^Stanford University.*$", re.I), "Stanford University"),
]

# Extra co-authors that appear on the published ASTRO 2026 listing but not in the CV entry (E191)
EXTRA_ROWS = [
    dict(key="Yu K", name="Yu, K.", org="Johns Hopkins University", conf="low",
         optional="[CONFIRM] named on the ASTRO 2026 listing of this abstract but not in the CV entry",
         url="https://amportal.astro.org/sessions/pqa-04-22880/cumulative-eqd2-esophageal-dose-and-toxicity-after-thoracic-reirradiation-114345",
         entries=[("E191", "Abstract", "Cumulative EQD2 Esophageal Dose and Severe Toxicity After Thoracic Reirradiation", "ASTRO Annual Conference", "09/2026")],
         notes="Not in the CV author list; added from the ASTRO 2026 program listing. Confirm final author list."),
    dict(key="Lin CT", name="Lin, C. T.", org="Johns Hopkins University", conf="low",
         optional="[CONFIRM] named on the ASTRO 2026 listing of this abstract but not in the CV entry",
         url="https://amportal.astro.org/sessions/pqa-04-22880/cumulative-eqd2-esophageal-dose-and-toxicity-after-thoracic-reirradiation-114345",
         entries=[("E191", "Abstract", "Cumulative EQD2 Esophageal Dose and Severe Toxicity After Thoracic Reirradiation", "ASTRO Annual Conference", "09/2026")],
         notes="Not in the CV author list; added from the ASTRO 2026 program listing. Confirm final author list."),
]

CV_TYPOS = [
    ("E122", "'Write J.' is Jean L. Wright (printed correctly as 'Wright J' elsewhere)."),
    ("E123", "'Arcot R.' and 'Rajasekar A.' are the same person, Arcot Rajasekar (given/family names swapped in one source)."),
    ("E127", "'Shan K. Y.' is Ketan Y. Shah (PubMed: 'Shah, Ketan Y')."),
    ("E131", "'Moncio A' is Alexander Moncion."),
    ("E179", "'Cen X' is Xuguang (Scott) Chen, who appears as 'Chen X' on E190."),
    ("E192", "'Krupein A' is most likely Andrew J. Krupien (UCLA); 'Maxin P' is most likely Peter G. Maxim (UC Irvine). Neither spelling exists in PubMed."),
    ("E191", "The ASTRO 2026 listing prints the first author as 'C. D. S. Smith', adds K. Yu and C. T. Lin, and omits Hagan, Gebre, Obaideen and Negassa. Check the CV entry against the final abstract."),
]

# ---------------------------------------------------------------- helpers
DEGREE_CUT = re.compile(r",.*$")

def clean_full(full):
    s = re.sub(r"\([^)]*\)", "", full)          # parentheticals
    s = re.sub(r"[\"“”][^\"“”]*[\"“”]", "", s)   # quoted nicknames
    s = DEGREE_CUT.sub("", s)                   # degrees after first comma
    return re.sub(r"\s+", " ", s).strip()

def display_name(key, full):
    if key in NAME_OVERRIDE:
        return NAME_OVERRIDE[key]
    parts = key.split(" ")
    surname = " ".join(parts[:-1]) if len(parts) > 1 else key
    s = clean_full(full)
    if not s or s.upper().startswith("UNKNOWN"):
        return f"{surname}, {parts[-1] if len(parts) > 1 else ''}".strip(", ")
    toks = s.split(" ")
    n = len(surname.split(" "))
    idx = None
    for i in range(len(toks) - n, -1, -1):
        if " ".join(toks[i:i + n]).lower() == surname.lower():
            idx = i
            break
    if idx is None:
        raise ValueError(f"surname {surname!r} not found in {full!r} for key {key!r}")
    sur = " ".join(toks[idx:idx + n])
    given = " ".join(toks[:idx] + toks[idx + n:])
    return f"{sur}, {given}".strip(", ")

def split_org(org):
    """-> (organization, department/detail) from the free-text org cell."""
    s = org.strip()
    detail = ""
    m = re.search(r"\s\(", s)
    if m:
        head, rest = s[:m.start()], s[m.start():]
        inner = re.match(r"\s\(([^)]*)\)", rest)
        detail = inner.group(1) if inner else ""
        s = head
    for sep in (" / ", " — ", " - possibly"):
        if sep in s:
            s = s.split(sep)[0]
    m2 = re.search(r",\s+(School of|Department of|Division of|Radiation Oncology|Dept)", s)
    if m2:
        detail = detail or s[m2.end() - len(m2.group(1)):]
        s = s[:m2.start()]
    s = s.strip(" ,;")
    for pat, repl in ORG_NORMALIZE:
        if pat.match(s):
            s = repl
            break
    # trim detail to the department phrase only
    d = re.sub(r"^(UC San Diego School of Medicine|UCI School of Medicine|Michigan Medicine|Keck School of Medicine of USC|UNC School of Medicine),\s*", "", detail)
    d = re.split(r"[;,]\s*(?:Houston|Ann Arbor|New York|Vienna|Stockholm|Philadelphia|Los Angeles|Palo Alto|Raleigh|New Delhi|Richmond|Groningen|Detroit|Charleston|Kansas City|Springfield|Lowell|Parkersburg|Long Beach)", d)[0]
    return s, d.strip(" ,;")

# ---------------------------------------------------------------- build Table 4
people = roster()
rows = []
for key in sorted(people):
    if key in DROP:
        continue
    info = lookup.get(key) or next((lookup[v] for v in sorted(people[key]["variants"]) if v in lookup), None)
    if info is None:
        raise KeyError(f"no lookup row for {key}")
    name = display_name(key, info["full"])
    if key in ORG_OVERRIDE:
        org, note = ORG_OVERRIDE[key]
        dept = ""
        flag = True
    else:
        org, dept = split_org(info["org"])
        note = ""
        flag = info["conf"] in ("medium", "low")
        if org.upper().startswith("UNKNOWN"):
            raise ValueError(f"UNKNOWN org without override: {key}")
    if key in ORG_FIX:
        org, dept = ORG_FIX[key]
    dept = DETAIL_OVERRIDE.get(key, dept)
    dept = re.sub(r"^School of Medicine,\s*", "", dept)
    if dept in DETAIL_DROP:
        dept = ""
    opt_parts = []
    if dept:
        opt_parts.append(dept)
    if flag:
        reason = note or FLAG_REASON.get(key)
        if not reason:
            raise ValueError(f"flagged row without an explicit reason: {key} ({info['conf']})")
        opt_parts.append("[CONFIRM] " + reason)
    optional = "; ".join(opt_parts)
    ents = people[key]["entries"]
    variants = sorted(people[key]["variants"])
    if key in [v for v in DROP.values()]:
        variants = sorted(set(variants) | {k for k, v in DROP.items() if v == key})
        ents = ents + people[[k for k, v in DROP.items() if v == key][0]]["entries"]
    rows.append(dict(key=key, variants=variants, name=name, org=org, optional=optional, url=info["url"],
                     conf=("low" if key in ORG_OVERRIDE and info["conf"] != "high" else info["conf"]),
                     notes=info["notes"], entries=sorted(set(ents), key=lambda e: e[0]), flag=flag))
for extra in EXTRA_ROWS:
    rows.append(dict(key=extra["key"], variants=[], name=extra["name"], org=extra["org"], optional=extra["optional"],
                     url=extra["url"], conf=extra["conf"], notes=extra["notes"], entries=extra["entries"], flag=True))
rows.sort(key=lambda r: r["name"].lower())

# ---------------------------------------------------------------- other tables
table1 = [
    dict(name="Anderson, Brian M.", org="University of California, San Diego",
         optional="Clinical Assistant Professor, Department of Radiation Medicine and Applied Sciences (primary)", last=""),
    dict(name="", org="University of North Carolina at Chapel Hill",
         optional="Adjunct Assistant Professor, Department of Radiation Oncology (2025-present)", last=""),
]
table2 = []
table3 = [
    dict(type="G", name="Brock, Kristy K.", org="The University of Texas MD Anderson Cancer Center",
         optional="Department of Imaging Physics; Ph.D. advisor (2017-2021)", last=""),
]
table5 = [
    dict(name="Boone, John M.", relationship="Editor-in-Chief (Brian is an Associate Editor, 2023-present)",
         org="University of California, Davis", optional="Medical Physics (AAPM/Wiley); EIC 2021-2026", last=""),
]

data = dict(reference_date=REF_DATE, table1=table1, table2=table2, table3=table3, table4=rows, table5=table5)
json.dump(data, open(OUT_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

# ---------------------------------------------------------------- source list
ent_by_id = {e[0]: e for e in ENTRIES}
L = []
L.append("# COA source list — Brian M. Anderson\n")
L.append(f"Reference (assumed submission) date: **{REF_DATE}**. Windows: Table 1 = 12 months, Table 4 = 48 months "
         "(publications dated 10/2022 or later), Table 5 = 24 months. See README.md.\n")
L.append("Source of truth for names and dates: `Clincal_CV/BMAnderson CV.docx`, Publications section, extracted with python-docx "
         "on 2026-10-27. Current organizations were looked up on 2026-10-27 (faculty pages, 2025-2026 PubMed/Europe PMC bylines, ORCID, "
         "conference listings); the URL in each row is the page that supports the organization entered.\n")
L.append("## Row counts\n")
n_flag = sum(1 for r in rows if r["flag"])
L.append("| Table | Rows | Notes |\n|---|---|---|")
L.append(f"| 1 — Your affiliations | {len(table1)} | UC San Diego (primary), UNC Chapel Hill (adjunct) |")
L.append(f"| 2 — Personal/family/business | {len(table2)} | left empty — Brian must answer |")
L.append(f"| 3 — Advisors/advisees | {len(table3)} | G = Kristy K. Brock; no T (thesis advisees) found in the CV |")
L.append(f"| 4 — Co-authors/collaborators | {len(rows)} | all type A; {n_flag} rows carry [CONFIRM]; 0 type C rows (see open questions) |")
L.append(f"| 5 — Co-editors | {len(table5)} | Editor-in-Chief of Medical Physics; co-editors unknown (flagged) |\n")

L.append("## Qualifying CV entries (Publications section, dated 10/2022 or later)\n")
L.append("| ID | Kind | Title (short) | Venue | Date |\n|---|---|---|---|---|")
for e in ENTRIES:
    L.append(f"| {e[0]} | {e[1]} | {e[2]} | {e[3]} | {e[4]} |")
L.append("")
L.append("E119 (book chapter) is undated in the CV; IOP Publishing's book page gives November 2025 "
         "(https://iopscience.iop.org/book/edit/978-0-7503-6120-0), chapter 9 \"Artificial intelligence-based image registration and "
         "segmentation\" by Brian M Anderson and Kristy K Brock. The 09/2022 and earlier entries (Woodland SASHIMI 09/2022, EPIDEEP 07/2022, "
         "Woodland AAPM 07/2022, etc.) fall outside the 48-month window and are excluded. E135 and E136 (both 10/2022) sit on the boundary and are included.\n")

L.append("## Table 1 — Your organizational affiliations (last 12 months)\n")
L.append("| Name | Organization | Optional | Source |\n|---|---|---|---|")
L.append("| Anderson, Brian M. | University of California, San Diego | Clinical Assistant Professor, Dept. of Radiation Medicine and Applied Sciences (primary) | CV header; biosketch 'Positions and Employment' (2025 – Present) |")
L.append("|  | University of North Carolina at Chapel Hill | Adjunct Assistant Professor, Dept. of Radiation Oncology (2025-present) | Biosketch_BMA.docx 'Positions and Employment': '2025 – Present Adjunct Assistant Professor, University of North Carolina, Chapel Hill' (not stated in the CV itself) |\n")

L.append("## Table 2 — Personal, family, or business relationships\n")
L.append("Left empty. Only Brian can answer this; nothing in the CV speaks to it. **[CONFIRM]**\n")

L.append("## Table 3 — Advisors and advisees\n")
L.append("| Type | Name | Organization | Source |\n|---|---|---|---|")
L.append("| G | Brock, Kristy K. | The University of Texas MD Anderson Cancer Center | Ph.D. advisor: CV Education (Ph.D. 2017–2021, MD Anderson) and dissertation; current position https://faculty.mdanderson.org/profiles/kristy_brock.html |")
L.append("")
L.append("No Ph.D. thesis advisees (T) appear anywhere in the CV. **[CONFIRM]** whether Brian has served as thesis advisor to any Ph.D. student. "
         "Master's advisor (Laurence Court) and postdoctoral sponsors are not requested by the current template.\n")

L.append("## Table 4 — Co-authors (A)\n")
L.append("Every person below co-authored at least one qualifying entry. Name is entered as 'Surname, Given names' where the given name "
         "could be established; otherwise as printed in the CV. 'CV key' shows the spelling(s) used in the CV.\n")
L.append("| # | Name entered | CV key(s) | Qualifying entries | Organization entered | Confidence | Affiliation source | Notes |")
L.append("|---|---|---|---|---|---|---|---|")
for i, r in enumerate(rows, 1):
    ents = "; ".join(f"{e[0]} ({e[1]}, {e[4]})" for e in r["entries"])
    keys = ", ".join(r["variants"]) if r["variants"] else "(not in CV)"
    note = r["notes"].replace("|", "/")
    if r["flag"]:
        note = "**[CONFIRM]** " + note
    L.append(f"| {i} | {r['name']} | {keys} | {ents} | {r['org']} | {r['conf']} | {r['url']} | {note} |")
L.append("")
L.append("### Collaborators (C)\n")
L.append("No C rows were entered. Every named collaborator in the brief (AHRQ R01 team at UNC, Open RT Structures / TG-263 co-authors, "
         "BigROC co-authors) already appears above as a co-author (A) on a product dated 10/2022 or later, and NSF lists each person once. "
         "AAPM committee co-members (Machine Intelligence Subcommittee, QUANTEC 2, Board of Associate Editors) are not listed because no "
         "co-authored product with them appears in the window. **[CONFIRM]**: add as C anyone Brian collaborated with on a project since "
         "10/2022 without a co-authored product (e.g., AHRQ project staff, Keras/TensorFlow maintainers who reviewed the reconstruct_patches "
         "contribution, COVER-ALL trial staff).\n")

L.append("## Table 5 — Co-editors (last 24 months)\n")
L.append("| Name | Relationship | Organization | Source |\n|---|---|---|---|")
L.append("| Boone, John M. | Editor-in-Chief of Medical Physics (Brian: Associate Editor 2023–present; member, AAPM Medical Physics Board of Associate Editors 2023–present) | University of California, Davis | AAPM release 'Paul K Appointed EIC July 2026' (https://w3.aapm.org/media/releases/Paul_K_Appointed_EIC_July_2026_Final.pdf): Kinahan succeeds John Boone, EIC 2021–2026, effective 1/1/2027; https://medphys.org/ lists John M. Boone as Editor-in-Chief |")
L.append("")
L.append("**[CONFIRM]**: Deputy/section editors Brian worked with directly as Associate Editor in the last 24 months are not derivable from the CV "
         "(medphys.org lists Stanley H. Benedict as Editor (Therapy)). Add any such editor; do not add editors or reviewers Brian dealt with only "
         "as a submitting author. If submission slips to 2027, replace Boone with Paul E. Kinahan (University of Washington).\n")

L.append("## CV transcription issues noticed (not edited — CV is read-only for this task)\n")
for eid, msg in CV_TYPOS:
    L.append(f"- **{eid}** ({ent_by_id[eid][1]}, {ent_by_id[eid][4]}): {msg}")
L.append("")

L.append("## Open questions for Brian\n")
L.append("1. Table 2: any personal, family, or business relationships to declare?")
L.append("2. Table 3: any Ph.D. thesis advisees?")
L.append(f"3. Table 4: {n_flag} rows marked [CONFIRM] (see the Notes column) — mostly trainees and staff who have left their 2022–2025 institutions, "
         "plus identity ambiguities (Chen S, Friend D, Dong F, Patel S, Hsu D, Kim S) and the two co-authors added from the ASTRO 2026 listing (Yu K, Lin CT).")
L.append("4. Table 4: any collaborators (type C) without a co-authored product since 10/2022?")
L.append("5. Table 5: which Medical Physics editors has Brian worked with directly in the last 24 months?")
L.append("6. Boundary: E135 and E136 are dated 10/2022; with a 10/27/2026 submission the window starts 10/27/2022, so these two papers "
         "(and co-authors who appear only on them: Weaver, Dennison, Awais, Chiang, Yang WT, Hanash, McCulloch, Cazoulat, Svensson, Gryshkevych, "
         "Kirimli, De, Mathew, Zaid, Elganainy, Peterson, Balter) could be dropped if the papers' actual publication dates are before 10/27/2022. "
         "They are included to err on the side of completeness.")
L.append("")
open(OUT_MD, "w", encoding="utf-8").write("\n".join(L))
print(f"table4 rows: {len(rows)}  flagged: {n_flag}")
for r in rows:
    print(f"{'*' if r['flag'] else ' '} {r['name']:34s} | {r['org'][:55]:55s} | {r['optional'][:70]}")
