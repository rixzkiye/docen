#!/usr/bin/env python3
"""Author the python-docx (independent writer) input fixtures.

Usage: python3 packages/docx/tests/oracle/author-foreign.py

Regenerates `fixtures/hf-foreign/input.docx`. The file is checked in so the
harness runs without python-docx installed; rerun this script to refresh it.
"""
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
OUT = HERE / "fixtures" / "hf-foreign" / "input.docx"


def set_pg_num(section, fmt=None, start=None, chap_style=None, chap_sep=None):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        pgSz = sectPr.find(qn("w:pgSz"))
        if pgSz is not None:
            pgSz.addprevious(pg)
        else:
            sectPr.append(pg)
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))
    if chap_style is not None:
        pg.set(qn("w:chapStyle"), str(chap_style))
    if chap_sep:
        pg.set(qn("w:chapSep"), chap_sep)


def add_field(paragraph, instr):
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr_el = OxmlElement("w:instrText")
    instr_el.set(qn("xml:space"), "preserve")
    instr_el.text = instr
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    paragraph.add_run()._r.append(begin)
    paragraph.add_run()._r.append(instr_el)
    paragraph.add_run()._r.append(end)


def fill(doc, tag, n):
    for i in range(n):
        p = doc.add_paragraph(f"{tag}-{i + 1}")
        p.paragraph_format.space_after = 0


def main():
    doc = Document()
    first = doc.sections[0]
    first.different_first_page_header_footer = True
    first.first_page_header.paragraphs[0].text = "FOREIGN-FIRST"
    first.header.paragraphs[0].text = "FOREIGN-DEFAULT"
    first.footer.paragraphs[0].text = "PAGE "
    add_field(first.footer.paragraphs[0], " PAGE ")
    set_pg_num(first, fmt="lowerRoman", start=1, chap_style=1, chap_sep="hyphen")
    fill(doc, "FGN1", 30)

    second = doc.add_section(WD_SECTION.NEW_PAGE)
    second.different_first_page_header_footer = False
    second.header.is_linked_to_previous = False
    second.header.paragraphs[0].text = "FOREIGN-SECOND"
    second.footer.is_linked_to_previous = False
    second.footer.paragraphs[0].text = "PAGE "
    add_field(second.footer.paragraphs[0], " PAGE ")
    set_pg_num(second, fmt="decimal", start=1)
    fill(doc, "FGN2", 60)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()
