#!/usr/bin/env python3
"""DOCX field/structural extractor for the toc-oracles fixture matrix.

Prints one JSON report:
  simple       [[instr, text], ...]   w:fldSimple fields, document order
  complex      [[instr, text], ...]   complex fields, document order (nested
                                      fields' text folds into the parent)
  instrTexts   ["...", ...]           every w:instrText value
  tocEntries   [{style, text, page}]  TOC/TOC-figure entry paragraphs, text
                                      split at the tab (page = trailing run)
  paragraphs   ["...", ...]           body paragraph texts
  parts        {"word/header1.xml": {simple, instrTexts, paragraphs}, ...}

Usage: field_extract.py <file.docx> [...]
"""
import json
import re
import sys
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag):
    return "{%s}%s" % (W, tag)


def tag_of(el):
    return etree.QName(el).localname


def text_of(el):
    return "".join(el.itertext())


def simple_fields(root):
    out = []
    for fld in root.iter(qn("fldSimple")):
        instr = fld.get(qn("instr")) or ""
        value = "".join(t.text or "" for t in fld.iter(qn("t")))
        out.append([instr.strip(), value])
    return out


def complex_fields(root):
    out = []
    stack = []
    for el in root.iter():
        tag = tag_of(el)
        if tag == "fldChar":
            kind = el.get(qn("fldCharType"))
            if kind == "begin":
                stack.append({"instr": "", "text": "", "sep": False})
            elif kind == "separate":
                if stack:
                    stack[-1]["sep"] = True
            elif kind == "end":
                if stack:
                    field = stack.pop()
                    out.append([field["instr"].strip(), field["text"]])
                    if stack and field["sep"]:
                        stack[-1]["text"] += field["text"]
            continue
        if tag == "instrText" and stack:
            stack[-1]["instr"] += el.text or ""
        elif tag == "t" and stack and stack[-1]["sep"]:
            stack[-1]["text"] += el.text or ""
    return out


ENTRY_STYLE = re.compile(r"^(toc\d*|tableoffigures|contents\d*)$", re.I)


def toc_entries(root):
    out = []
    for p in root.iter(qn("p")):
        style = None
        ppr = p.find(qn("pPr"))
        if ppr is not None:
            pstyle = ppr.find(qn("pStyle"))
            if pstyle is not None:
                style = pstyle.get(qn("val"))
        if not style or not ENTRY_STYLE.match(style):
            continue
        # Entry text = concatenated w:t runs; the page number follows the tab.
        pieces = []
        after_tab = False
        for el in p.iter():
            tag = tag_of(el)
            # w:tab appears both as a run-content tab (no w:val) and as a tab
            # stop definition inside w:pPr (w:val/pos) — only the former splits.
            if tag == "tab" and el.get(qn("val")) is None:
                after_tab = True
            elif tag == "t":
                pieces.append((after_tab, el.text or ""))
        text = "".join(value for after, value in pieces if not after)
        page = "".join(value for after, value in pieces if after)
        out.append({"style": style, "text": text, "page": page})
    return out


def paragraphs(root):
    return [text_of(p) for p in root.iter(qn("p"))]


def root_of(data):
    return etree.fromstring(data)


def part_report(data):
    root = root_of(data)
    return {
        "simple": simple_fields(root),
        "instrTexts": [el.text or "" for el in root.iter(qn("instrText"))],
        "paragraphs": paragraphs(root),
    }


def report(path):
    with zipfile.ZipFile(path) as archive:
        names = [n for n in archive.namelist() if re.match(r"word/(header|footer)\d*\.xml$", n)]
        xml = archive.read("word/document.xml")
        root = root_of(xml)
        out = {
            "simple": simple_fields(root),
            "complex": complex_fields(root),
            "instrTexts": [el.text or "" for el in root.iter(qn("instrText"))],
            "tocEntries": toc_entries(root),
            "paragraphs": paragraphs(root),
            "parts": {name: part_report(archive.read(name)) for name in sorted(names)},
        }
        return out


if __name__ == "__main__":
    result = {path: report(path) for path in sys.argv[1:]}
    print(json.dumps(result, ensure_ascii=False, indent=1))
