#!/usr/bin/env python3
"""Apply dark theme styling to all slides in a PowerPoint deck."""

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.oxml.ns import qn

from theme import (
    BODY_TEXT,
    CISCO_BLUE,
    COLOR_MAP,
    DARK_BANNER,
    DARK_NAVY,
    DARK_ROW,
    FOOTER_TEXT,
    WHITE,
)


def _set_solid_fill(element, rgb: RGBColor) -> None:
    """Set solid fill on a fill object."""
    element.solid()
    element.fore_color.rgb = rgb


def _set_text_color(run_or_rpr, rgb: RGBColor) -> None:
    """Set font color on a run or defRPr element."""
    if hasattr(run_or_rpr, "font"):
        run_or_rpr.font.color.rgb = rgb
    else:
        from lxml import etree

        solid = run_or_rpr.find(qn("a:solidFill"))
        if solid is None:
            solid = etree.SubElement(run_or_rpr, qn("a:solidFill"))
        srgb = solid.find(qn("a:srgbClr"))
        if srgb is None:
            srgb = etree.SubElement(solid, qn("a:srgbClr"))
        val = f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
        srgb.set("val", val)


def _recolor_text_frame(text_frame, dark_body: RGBColor = BODY_TEXT) -> None:
    """Recolor all text in a text frame for dark backgrounds."""
    for para in text_frame.paragraphs:
        # Paragraph-level default run properties
        pPr = para._p.find(qn("a:pPr"))
        if pPr is not None:
            defRPr = pPr.find(qn("a:defRPr"))
            if defRPr is not None:
                solid = defRPr.find(qn("a:solidFill"))
                if solid is not None:
                    srgb = solid.find(qn("a:srgbClr"))
                    if srgb is not None:
                        val = srgb.get("val", "").upper()
                        if val == "333333":
                            _set_text_color(defRPr, dark_body)
                        elif val == "666666":
                            _set_text_color(defRPr, FOOTER_TEXT)
                        elif val == "007BC7":
                            pass  # keep accent headings
                        elif val == "FFFFFF":
                            pass  # keep white text

        for run in para.runs:
            try:
                if run.font.color.type is not None:
                    rgb = run.font.color.rgb
                    if rgb == RGBColor(0x33, 0x33, 0x33):
                        run.font.color.rgb = dark_body
                    elif rgb == RGBColor(0x66, 0x66, 0x66):
                        run.font.color.rgb = FOOTER_TEXT
            except (AttributeError, TypeError):
                pass


def _recolor_table(table) -> None:
    """Apply dark theme to table cells."""
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            fill = cell.fill
            if fill.type == MSO_FILL.SOLID:
                try:
                    rgb = fill.fore_color.rgb
                    if rgb == RGBColor(0xF5, 0xF5, 0xF5):
                        _set_solid_fill(fill, DARK_ROW)
                    elif rgb == RGBColor(0x00, 0x7B, 0xC7):
                        pass  # keep header row
                except (AttributeError, TypeError):
                    pass
            elif row_idx > 0:
                # Unfilled data rows: ensure light text on dark slide
                pass

            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        if run.font.color.type is not None:
                            rgb = run.font.color.rgb
                            if rgb == RGBColor(0x33, 0x33, 0x33):
                                run.font.color.rgb = BODY_TEXT
                            elif rgb == RGBColor(0xFF, 0xFF, 0xFF):
                                pass  # header text stays white
                    except (AttributeError, TypeError):
                        pass

                pPr = para._p.find(qn("a:pPr"))
                if pPr is not None:
                    defRPr = pPr.find(qn("a:defRPr"))
                    if defRPr is not None:
                        solid = defRPr.find(qn("a:solidFill"))
                        if solid is not None:
                            srgb = solid.find(qn("a:srgbClr"))
                            if srgb is not None and srgb.get("val", "").upper() == "333333":
                                _set_text_color(defRPr, BODY_TEXT)


def _set_slide_background(slide, color: RGBColor = DARK_NAVY) -> None:
    """Set solid dark background on a slide."""
    fill = slide.background.fill
    _set_solid_fill(fill, color)


def _recolor_shape_fills(shape) -> None:
    """Recolor shape fills (subtitle banners, etc.)."""
    if not hasattr(shape, "fill"):
        return
    fill = shape.fill
    if fill.type == MSO_FILL.SOLID:
        try:
            rgb = fill.fore_color.rgb
            if rgb == RGBColor(0xCC, 0xE8, 0xF7):
                _set_solid_fill(fill, DARK_BANNER)
        except (AttributeError, TypeError):
            pass


def apply_dark_theme_to_presentation(prs: Presentation) -> None:
    """Apply dark theme to every slide in a presentation."""
    for slide in prs.slides:
        _set_slide_background(slide)

        for shape in slide.shapes:
            _recolor_shape_fills(shape)

            if shape.has_text_frame:
                _recolor_text_frame(shape.text_frame)

            if shape.has_table:
                _recolor_table(shape.table)


def patch_pptx_xml(pptx_path: Path) -> None:
    """Post-process PPTX XML for colors not reachable via python-pptx API."""
    tmp_path = pptx_path.with_suffix(".tmp.pptx")
    shutil.copy2(pptx_path, tmp_path)

    with zipfile.ZipFile(tmp_path, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}

    for name, data in entries.items():
        if not (name.endswith(".xml") or name.endswith(".rels")):
            continue
        text = data.decode("utf-8")
        for old, new in COLOR_MAP.items():
            text = text.replace(f'srgbClr val="{old}"', f'srgbClr val="{new}"')
            text = text.replace(f'srgbClr val="{old.lower()}"', f'srgbClr val="{new}"')
        entries[name] = text.encode("utf-8")

    # Dark theme color scheme
    theme_key = "ppt/theme/theme1.xml"
    if theme_key in entries:
        theme = entries[theme_key].decode("utf-8")
        theme = theme.replace('name="Office Theme"', 'name="Cisco Dark Theme"')
        theme = theme.replace('name="Office"', 'name="Cisco Dark"')
        theme = re.sub(
            r"<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>",
            '<a:lt1><a:srgbClr val="1A2B4A"/></a:lt1>',
            theme,
        )
        theme = re.sub(
            r"<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>",
            '<a:dk1><a:srgbClr val="E0E0E0"/></a:dk1>',
            theme,
        )
        theme = re.sub(
            r"<a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>",
            '<a:lt2><a:srgbClr val="2D3A52"/></a:lt2>',
            theme,
        )
        theme = re.sub(
            r"<a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>",
            '<a:dk2><a:srgbClr val="007BC7"/></a:dk2>',
            theme,
        )
        theme = re.sub(
            r"<a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>",
            '<a:accent1><a:srgbClr val="007BC7"/></a:accent1>',
            theme,
        )
        entries[theme_key] = theme.encode("utf-8")

    with zipfile.ZipFile(pptx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries.items():
            zout.writestr(name, data)

    tmp_path.unlink()


def apply_dark_theme(input_path: Path, output_path: Path | None = None) -> Path:
    """Apply dark theme to a PPTX file and return the output path."""
    output = output_path or input_path
    working = input_path if output_path is None else output_path

    if output_path and output_path != input_path:
        shutil.copy2(input_path, working)

    prs = Presentation(str(working))
    apply_dark_theme_to_presentation(prs)
    prs.save(str(working))
    patch_pptx_xml(working)
    return working


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Apply dark theme to PowerPoint decks")
    parser.add_argument("input", type=Path, help="Input .pptx file")
    parser.add_argument("-o", "--output", type=Path, help="Output .pptx file (default: in-place)")
    args = parser.parse_args(argv)

    if not args.input.exists():
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        return 1

    result = apply_dark_theme(args.input, args.output)
    print(f"Dark theme applied: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
