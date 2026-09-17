#!/usr/bin/env python3
"""fontTools validator for R6 subset output.

Reads a JSON config path as argv[1]:
  {
    "original": "<ttf path>",
    "subset": "<ttf path>",
    "checks": [{"cp": <codepoint>, "subsetGid": <int>, "advance": <int>}]
  }

Exits non-zero with a report when the subset does not parse, cmap/post are
stale, an outline is malformed, a composite component is missing, an advance
drifted, or a save round-trip fails.
"""

import json
import os
import sys
import tempfile

from fontTools.ttLib import TTFont


def main() -> int:
    with open(sys.argv[1], "r", encoding="utf-8") as fh:
        cfg = json.load(fh)

    errors = []
    advances = {}
    try:
        original = TTFont(cfg["original"], fontNumber=0)
        subset = TTFont(cfg["subset"], fontNumber=0)
    except Exception as exc:  # noqa: BLE001 - report parse failure verbatim
        print(json.dumps({"errors": [f"parse failed: {exc!r}"]}))
        return 1

    try:
        orig_cmap = original.getBestCmap()
        sub_cmap = subset.getBestCmap()
        order = subset.getGlyphOrder()
        glyf = subset["glyf"]
        hmtx_orig = original["hmtx"]
        hmtx_sub = subset["hmtx"]

        if len(order) != subset["maxp"].numGlyphs:
            errors.append(
                f"maxp.numGlyphs {subset['maxp'].numGlyphs} != glyph order {len(order)}"
            )

        for check in cfg["checks"]:
            cp = check["cp"]
            gid_name = sub_cmap.get(cp)
            if gid_name is None:
                errors.append(f"U+{cp:04X} missing from subset cmap")
                continue
            gid = order.index(gid_name)
            if gid != check["subsetGid"]:
                errors.append(
                    f"U+{cp:04X} maps to {gid}, expected subsetGid {check['subsetGid']}"
                )
            orig_name = orig_cmap.get(cp)
            if orig_name is not None:
                advances[str(cp)] = hmtx_orig[orig_name][0]
                if hmtx_sub[gid_name][0] != hmtx_orig[orig_name][0]:
                    errors.append(
                        f"U+{cp:04X} advance {hmtx_sub[gid_name][0]} != original {hmtx_orig[orig_name][0]}"
                    )
            glyph = glyf[gid_name]
            try:
                if glyph.isComposite():
                    for component in glyph.components:
                        if component.glyphName not in order:
                            errors.append(
                                f"composite {gid_name} references missing {component.glyphName}"
                            )
                else:
                    glyph.getCoordinates(glyf)
            except Exception as exc:  # noqa: BLE001 - malformed outline
                errors.append(f"outline {gid_name} malformed: {exc!r}")

        if float(subset["post"].formatType) != 3.0:
            errors.append(f"post formatType {subset['post'].formatType} != 3.0")

        # A full save exercises every table's compile path.
        with tempfile.TemporaryDirectory() as tmp:
            out = os.path.join(tmp, "roundtrip.ttf")
            subset.save(out)
            reparsed = TTFont(out)
            if len(reparsed.getGlyphOrder()) != len(order):
                errors.append("save round-trip changed glyph order length")
    except Exception as exc:  # noqa: BLE001 - report validation failure verbatim
        errors.append(f"validation crashed: {exc!r}")

    print(json.dumps({"errors": errors, "glyphCount": len(subset.getGlyphOrder()), "advances": advances}))
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
