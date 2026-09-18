#!/usr/bin/env python3
"""Independent DOCX structure oracle.

Prints one JSON object per input file with structural signals extracted by
python-docx (package open) and lxml (raw OOXML): section references and
properties, note parts (ids, separators, refs, images), header/footer parts,
dangling relationships and content-type gaps. run.mjs evaluates the fixture
`python` checks against these signals.

Usage: python3 oracle.py <file.docx> [<file.docx> ...]
"""
from __future__ import annotations

import json
import sys
import zipfile
from urllib.parse import unquote

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def attr(el, name):
    return el.get(f"{W}{name}") if el is not None else None


def int_attr(el, name):
    value = attr(el, name)
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def child(el, name):
    return el.find(f"{W}{name}") if el is not None else None


def text_of(el) -> str:
    return "".join(el.itertext())


def note_props(el):
    props = {}
    if el is None:
        return props
    pos = child(el, "pos")
    if pos is not None:
        props["pos"] = attr(pos, "val")
    num_fmt = child(el, "numFmt")
    if num_fmt is not None:
        if attr(num_fmt, "val") is not None:
            props["formatType"] = attr(num_fmt, "val")
        if attr(num_fmt, "format") is not None:
            props["format"] = attr(num_fmt, "format")
    num_start = child(el, "numStart")
    if num_start is not None:
        props["numStart"] = int_attr(num_start, "val")
    num_restart = child(el, "numRestart")
    if num_restart is not None:
        props["numRestart"] = attr(num_restart, "val")
    return props


def section_signals(sect):
    refs = {"headerRefs": {}, "footerRefs": {}}
    for ref in sect.findall(f"{W}headerReference"):
        refs["headerRefs"][attr(ref, "type")] = ref.get(f"{R}id")
    for ref in sect.findall(f"{W}footerReference"):
        refs["footerRefs"][attr(ref, "type")] = ref.get(f"{R}id")
    pg = child(sect, "pgNumType")
    pg_num = {}
    if pg is not None:
        if attr(pg, "fmt"):
            pg_num["format"] = attr(pg, "fmt")
        if int_attr(pg, "start") is not None:
            pg_num["start"] = int_attr(pg, "start")
        if attr(pg, "chapSep"):
            pg_num["separator"] = attr(pg, "chapSep")
        if int_attr(pg, "chapStyle") is not None:
            pg_num["chapterStyle"] = int_attr(pg, "chapStyle")
    title = child(sect, "titlePg")
    no_endnote = child(sect, "noEndnote")
    sect_type = child(sect, "type")
    return {
        **refs,
        "titlePg": bool(title is not None and attr(title, "val") not in ("0", "false")),
        "pgNumType": pg_num,
        "footnotePr": note_props(child(sect, "footnotePr")),
        "endnotePr": note_props(child(sect, "endnotePr")),
        "noEndnote": (
            attr(no_endnote, "val") not in ("0", "false") if no_endnote is not None else None
        ),
        "type": attr(sect_type, "val") if sect_type is not None else None,
    }


def notes_signals(root, note_tag, part_name, rels):
    if root is None:
        return None
    ids, separators, notices, image_targets, texts = [], [], [], [], []
    for el in root.findall(f"{W}{note_tag}"):
        note_id = int_attr(el, "id")
        note_type = attr(el, "type")
        if note_type == "separator":
            separators.append(note_id)
        elif note_type == "continuationSeparator":
            separators.append(note_id)
        elif note_type == "continuationNotice":
            notices.append(note_id)
        else:
            ids.append(note_id)
        texts.append("".join(text_of(el).split()))
    ref_marks = len(root.findall(f".//{W}{note_tag}Ref"))
    for rel in rels.get(part_rels_key(part_name), []):
        if rel["kind"] == "image":
            image_targets.append(rel["target"])
    return {
        "ids": sorted(n for n in ids if n is not None),
        "separatorIds": sorted(n for n in separators if n is not None),
        "continuationNoticeIds": sorted(n for n in notices if n is not None),
        "refMarks": ref_marks,
        "imageTargets": image_targets,
        "texts": texts,
    }


def part_rels_key(part: str) -> str:
    """`word/footnotes.xml` → the rels map key `word/_rels/footnotes.xml`."""
    directory, _, name = part.rpartition("/")
    return f"{directory}/_rels/{name}"


def rels_of(zf, names):
    out = {}
    for name in names:
        if not name.endswith(".rels"):
            continue
        try:
            root = etree.fromstring(zf.read(name))
        except etree.XMLSyntaxError:
            continue
        entries = []
        for rel in root.findall(f"{REL}Relationship"):
            entries.append(
                {
                    "id": rel.get("Id"),
                    "type": rel.get("Type"),
                    "kind": (rel.get("Type") or "").rsplit("/", 1)[-1],
                    "target": rel.get("Target"),
                    "external": rel.get("TargetMode") == "External",
                }
            )
        out[name[:-5]] = entries
    return out


def owner_part(rels_key: str) -> str:
    """`word/_rels/header1.xml` → the part owning those rels; the package
    root rels (`_rels/`) owns the package itself."""
    if rels_key == "_rels/":
        return ""
    if "/_rels/" in rels_key:
        return rels_key.replace("/_rels/", "/")
    return rels_key


def resolve_target(owner: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = owner.rsplit("/", 1)[0] if "/" in owner else ""
    stack = []
    for part in (f"{base}/{target}" if base else target).split("/"):
        if part == "..":
            if stack:
                stack.pop()
        elif part not in ("", "."):
            stack.append(part)
    return "/".join(stack)


def inspect(path: str) -> dict:
    report = {"file": path, "ok": False, "errors": []}
    try:
        import docx

        document = docx.Document(path)
        report["pythonDocx"] = {
            "paragraphs": len(document.paragraphs),
            "tables": len(document.tables),
            "sections": len(document.sections),
            "inlineShapes": len(document.inline_shapes),
        }
    except Exception as exc:  # noqa: BLE001 - report, never crash the oracle
        report["errors"].append(f"python-docx: {exc}")

    try:
        zf = zipfile.ZipFile(path)
    except Exception as exc:  # noqa: BLE001
        report["errors"].append(f"zipfile: {exc}")
        return report

    names = zf.namelist()
    report["parts"] = names
    report["media"] = {
        "count": len([n for n in names if n.startswith("word/media/")]),
        "targets": sorted(n for n in names if n.startswith("word/media/")),
    }
    for name in names:
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        try:
            etree.fromstring(zf.read(name))
        except etree.XMLSyntaxError:
            report["errors"].append(f"xml parse: {name}")

    rels = rels_of(zf, names)
    report["danglingRels"] = []
    for rels_key, entries in rels.items():
        owner = owner_part(rels_key)
        for rel in entries:
            if rel["external"]:
                continue
            target = resolve_target(owner, rel["target"] or "")
            if unquote(target) not in names:
                report["danglingRels"].append({"owner": owner, **rel})

    defaults, overrides = {}, set()
    ct_root = etree.fromstring(zf.read("[Content_Types].xml"))
    for default in ct_root.findall(f"{CT}Default"):
        defaults[(default.get("Extension") or "").lower()] = default.get("ContentType")
    for override in ct_root.findall(f"{CT}Override"):
        overrides.add((override.get("PartName") or "").lstrip("/"))
    gaps = []
    for name in names:
        if name == "[Content_Types].xml" or name.endswith(".rels"):
            continue
        if name in overrides:
            continue
        if name.rsplit(".", 1)[-1].lower() in defaults:
            continue
        gaps.append(name)
    gaps.extend(f"override-without-part: {part}" for part in overrides if part not in names)
    report["contentTypeGaps"] = gaps

    document_xml = read_xml(zf, "word/document.xml")
    sections, footnote_refs, endnote_refs = [], [], []
    if document_xml is not None:
        for sect in document_xml.iter(f"{W}sectPr"):
            sections.append(section_signals(sect))
        footnote_refs = [int_attr(el, "id") for el in document_xml.iter(f"{W}footnoteReference")]
        endnote_refs = [int_attr(el, "id") for el in document_xml.iter(f"{W}endnoteReference")]
    report["document"] = {
        "sections": sections,
        "footnoteRefs": footnote_refs,
        "endnoteRefs": endnote_refs,
    }

    report["notes"] = {
        "footnotes": notes_signals(
            read_xml(zf, "word/footnotes.xml"), "footnote", "word/footnotes.xml", rels
        ),
        "endnotes": notes_signals(
            read_xml(zf, "word/endnotes.xml"), "endnote", "word/endnotes.xml", rels
        ),
    }

    settings = {}
    settings_xml = read_xml(zf, "word/settings.xml")
    if settings_xml is not None:
        even = settings_xml.find(f"{W}evenAndOddHeaders")
        settings["evenAndOddHeaders"] = even is not None and attr(even, "val") not in ("0", "false")
        settings["footnotePr"] = note_props(child(settings_xml, "footnotePr"))
        settings["endnotePr"] = note_props(child(settings_xml, "endnotePr"))
    report["settings"] = settings

    report["headers"] = part_signals(zf, names, rels, "header")
    report["footers"] = part_signals(zf, names, rels, "footer")

    report["ok"] = (
        not report["errors"] and not report["danglingRels"] and not report["contentTypeGaps"]
    )
    return report


def part_signals(zf, names, rels, kind):
    import re

    out = {}
    for name in names:
        match = re.fullmatch(rf"word/{kind}(\d+)\.xml", name)
        if not match:
            continue
        root = read_xml(zf, name)
        if root is None:
            continue
        images = [
            rel["target"]
            for rel in rels.get(part_rels_key(name), [])
            if rel["kind"] == "image"
        ]
        out[name.split("/")[-1].replace(".xml", "")] = {
            "texts": ["".join(text_of(p).split()) for p in root.findall(f"{W}p")],
            "images": images,
        }
    return out


def read_xml(zf, name):
    try:
        return etree.fromstring(zf.read(name))
    except KeyError:
        return None
    except etree.XMLSyntaxError:
        return None


def main() -> int:
    for path in sys.argv[1:]:
        print(json.dumps(inspect(path), ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
