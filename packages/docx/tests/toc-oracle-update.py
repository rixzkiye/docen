#!/usr/bin/env python3
"""Explicit LibreOffice field update for the toc-oracles oracle.

Connects to a soffice process listening on a UNO socket, loads the DOCX hidden,
runs an explicit update (all text fields + document indexes/TOC), then stores a
PDF and (optionally) a DOCX. This is the "user pressed Update All Fields /
Update Table" path — distinct from the plain load+convert oracle, which keeps
the document's cached results.

Usage: toc-oracle-update.py <in.docx> <out.pdf> [out.docx] [--port N]
"""

import os
import sys

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def main():
    argv = sys.argv[1:]
    port = 2002
    if "--port" in argv:
        index = argv.index("--port")
        port = int(argv[index + 1])
        del argv[index : index + 2]
    args = [a for a in argv if not a.startswith("--")]
    src, pdf_out = args[0], args[1]
    docx_out = args[2] if len(args) > 2 else None

    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local
    )
    ctx = resolver.resolve(
        f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
    )
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    url = uno.systemPathToFileUrl(os.path.abspath(src))
    doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))

    updated_fields = 0
    enum = doc.getTextFields().createEnumeration()
    while enum.hasMoreElements():
        field = enum.nextElement()
        try:
            field.update()
            updated_fields += 1
        except Exception:
            pass
    indexes = doc.getDocumentIndexes()
    index_count = indexes.getCount()
    for i in range(index_count):
        indexes.getByIndex(i).update()
    doc.refresh()
    doc.storeToURL(
        uno.systemPathToFileUrl(os.path.abspath(pdf_out)),
        (prop("FilterName", "writer_pdf_Export"),),
    )
    if docx_out:
        doc.storeToURL(
            uno.systemPathToFileUrl(os.path.abspath(docx_out)),
            (prop("FilterName", "MS Word 2007 XML"),),
        )
    print(
        "RESULT "
        + '{"updated_fields": %d, "indexes": %d}' % (updated_fields, index_count),
        flush=True,
    )
    doc.close(False)


if __name__ == "__main__":
    main()
