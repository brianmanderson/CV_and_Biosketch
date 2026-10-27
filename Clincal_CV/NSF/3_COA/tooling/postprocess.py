# -*- coding: utf-8 -*-
"""Restore template parts that openpyxl rewrites or drops, without touching cell data:
 - each xl/tables/tableN.xml is replaced by the ORIGINAL table XML (keeps altText accessibility
   extension and Excel revision ids) with only the ref / autoFilter ref updated to the new ranges;
 - the <cols> block of the sheet is replaced by the original one (restores width="0" on hidden columns);
 - the original (inactive) <sheetProtection .../> element is re-inserted after </sheetData>.
Usage: postprocess.py original.xlsx openpyxl_output.xlsx final.xlsx"""
import re, sys, zipfile

ORIG, SRC, DST = sys.argv[1:4]
zo, zs = zipfile.ZipFile(ORIG), zipfile.ZipFile(SRC)

def display_name(xml):
    return re.search(r'displayName="([^"]+)"', xml).group(1)

orig_tables = {}
for n in zo.namelist():
    if n.startswith("xl/tables/"):
        x = zo.read(n).decode("utf-8")
        orig_tables[display_name(x)] = x

sheet_src = zs.read("xl/worksheets/sheet1.xml").decode("utf-8")
sheet_orig = zo.read("xl/worksheets/sheet1.xml").decode("utf-8")

# cols block
cols_o = re.search(r"<cols>.*?</cols>", sheet_orig, re.S).group(0)
sheet_new, k = re.subn(r"<cols>.*?</cols>", lambda m: cols_o, sheet_src, count=1, flags=re.S)
assert k == 1
# sheetProtection (inactive in the template: no sheet="1")
prot = re.search(r"<sheetProtection[^>]*/>", sheet_orig)
if prot and "<sheetProtection" not in sheet_new:
    sheet_new = sheet_new.replace("</sheetData>", "</sheetData>" + prot.group(0), 1)


# headerFooter: openpyxl re-encodes the newline in the page header as the literal text _x000a_; restore the original block
hf_o = re.search(r"<headerFooter.*?</headerFooter>|<headerFooter[^>]*/>", sheet_orig, re.S)
if hf_o:
    sheet_new, k2 = re.subn(r"<headerFooter.*?</headerFooter>|<headerFooter[^>]*/>", lambda m: hf_o.group(0), sheet_new, count=1, flags=re.S)
    assert k2 == 1, "headerFooter block not found in openpyxl output"

with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as zd:
    for info in zs.infolist():
        n = info.filename
        if n == "xl/worksheets/sheet1.xml":
            data = sheet_new.encode("utf-8")
        elif n.startswith("xl/tables/"):
            new_xml = zs.read(n).decode("utf-8")
            name = display_name(new_xml)
            new_ref = re.search(r'<table [^>]*\bref="([^"]+)"', new_xml).group(1)
            x = orig_tables[name]
            x = re.sub(r'(<table [^>]*\bref=")[^"]+(")', lambda m: m.group(1) + new_ref + m.group(2), x, count=1)
            x = re.sub(r'(<autoFilter [^>]*\bref=")[^"]+(")', lambda m: m.group(1) + new_ref + m.group(2), x, count=1)
            print(f"{n}: {name} ref -> {new_ref}")
            data = x.encode("utf-8")
        else:
            data = zs.read(n)
        zd.writestr(info, data)
print("wrote", DST)
