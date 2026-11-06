#!/usr/bin/env python3
"""Mechanical checks for a lab meeting deck. Reading the rendered slides is still required.

An error means the file is broken or names someone it should not. A warning means a departure from the house style.

Usage:
    check_deck.py deck.pptx [--people Ann,Bob]    exit 0 clean or warnings only, 1 errors, 2 no such file
    check_deck.py deck.pptx --dump-numbers        every line with a digit, as the claim list for number_audit.tsv
"""

import argparse
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from deck_style import FONT, NO_STYLE

TOLERANCE_IN = 0.02
EMU_PER_IN = 914400
MAX_CHART_DPI = 260
MAX_TITLE_CHARS = 72  # 28 pt bold Times New Roman over the 12.33 in title box
DECIMAL_METRIC = re.compile(r"(?<![\w.])[+\-−±]0\.\d+(?![\d.]*%)|\bAUC\b[^0-9%]{0,12}0\.\d+(?![\d.]*%)")
IN_HOUSE_WORDS = re.compile(r"\barms?\b|5-model|cross-fitting", re.I)
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


@dataclass(frozen=True)
class Finding:
    level: str
    slide: int
    message: str


def _walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(shape.shapes)


def _text_frames(shape):
    if shape.has_text_frame:
        yield shape.text_frame
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                yield cell.text_frame


def _notes(slide):
    return slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""


def _is_warm(color):
    if color.type != MSO_COLOR_TYPE.RGB:
        return False
    r, _, b = color.rgb
    return r >= 150 and r - b >= 70


def _run_findings(slide, n):
    found = set()
    for shape in _walk(slide.shapes):
        for tf in _text_frames(shape):
            for p in tf.paragraphs:
                for run in p.runs:
                    if not run.text.strip():
                        continue
                    if run.font.name != FONT:
                        found.add(Finding("WARN", n, f"font {run.font.name or 'inherited from the theme'}, expected {FONT}: {run.text[:40]!r}"))
                    if _is_warm(run.font.color):
                        found.add(Finding("WARN", n, f"red or orange text: {run.text[:40]!r}"))
    return found


def _picture_finding(shape, n):
    dpi = shape.image.size[0] / (shape.width / EMU_PER_IN)
    if dpi > MAX_CHART_DPI:
        return Finding("WARN", n, f"image placed at {dpi:.0f} dpi; if it is a chart, draw it at slide size so its text is not shrunk")
    return None


def _shape_findings(slide, n, width, height):
    found = set()
    tolerance = TOLERANCE_IN * EMU_PER_IN
    for shape in _walk(slide.shapes):
        if shape._element.find(qn("p:style")) is not None:
            found.add(Finding("WARN", n, f"{shape.name} keeps the theme style block, which draws a shadow"))
        if getattr(shape, "has_table", False) and shape.has_table:
            props = shape.table._tbl.tblPr
            style = props.find(qn("a:tableStyleId"))
            if props.get("bandRow") == "1" or style is None or style.text != NO_STYLE:
                found.add(Finding("WARN", n, "table uses a banded or themed style; use deck_style.table"))
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            found.add(_picture_finding(shape, n))
        if shape.left is None or shape.top is None:
            continue
        if shape.left < -tolerance or shape.top < -tolerance or shape.left + shape.width > width + tolerance or shape.top + shape.height > height + tolerance:
            found.add(Finding("ERROR", n, f"{shape.name} runs off the slide"))
    return found - {None}


def _slide_lines(slide):
    for shape in _walk(slide.shapes):
        for tf in _text_frames(shape):
            for p in tf.paragraphs:
                if p.text.strip():
                    yield p.text


def _word_findings(slide, n, people):
    found = set()
    notes = _notes(slide)
    lines = list(_slide_lines(slide))
    title = slide.shapes.title.text_frame.text if slide.shapes.title is not None else ""
    if len(title) > MAX_TITLE_CHARS:
        found.add(Finding("WARN", n, f"title has {len(title)} characters, so it wraps onto the takeaway: {title[:50]!r}"))
    if not notes.strip():
        found.add(Finding("WARN", n, "no speaker notes"))
    for text in [*lines, notes]:
        if "—" in text:
            found.add(Finding("WARN", n, f"em dash: {text[:50]!r}"))
        for person in people:
            if re.search(rf"\b{re.escape(person)}\b", text):
                found.add(Finding("ERROR", n, f"names {person}, who is in the room: {text[:50]!r}"))
        hit = DECIMAL_METRIC.search(text)
        if hit:
            found.add(Finding("WARN", n, f"write this metric as a percent: {hit.group(0)!r} in {text[:50]!r}"))
    for text in lines:
        hit = IN_HOUSE_WORDS.search(text)
        if hit:
            found.add(Finding("WARN", n, f"in-house wording the audience may not follow: {hit.group(0)!r}"))
    return found


def _package_findings(path):
    found = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        types = etree.fromstring(z.read("[Content_Types].xml"))
        for override in types.findall(f"{CT_NS}Override"):
            part = override.get("PartName")
            if part.lstrip("/") not in names:
                found.append(Finding("ERROR", 0, f"content type declared for a missing part: {part}"))
        order = [child.tag for child in etree.fromstring(z.read("ppt/presentation.xml"))]
        notes_list, slide_list = f"{P_NS}notesMasterIdLst", f"{P_NS}sldIdLst"
        if notes_list in order and slide_list in order and order.index(notes_list) > order.index(slide_list):
            found.append(Finding("ERROR", 0, "notesMasterIdLst after sldIdLst; PowerPoint opens this with a Repair prompt"))
        for name in sorted(n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)):
            ids = [e.get("id") for e in etree.fromstring(z.read(name)).iter(f"{P_NS}cNvPr")]
            if len(ids) != len(set(ids)):
                found.append(Finding("ERROR", 0, f"duplicate shape id in {name}"))
    return found


def check(path, people=()):
    try:
        prs = Presentation(path)
    except Exception as error:  # a file python-pptx cannot read is itself the finding
        return [Finding("ERROR", 0, f"cannot open the deck: {error}")]
    found = _package_findings(path)
    slides = list(prs.slides)
    for n, slide in enumerate(slides, 1):
        found += sorted(
            _run_findings(slide, n) | _shape_findings(slide, n, prs.slide_width, prs.slide_height) | _word_findings(slide, n, people),
            key=lambda f: (f.level, f.message),
        )
    title_slides = sum(slide.slide_layout.name == "Title Slide" for slide in slides)
    if title_slides > 1:
        found.append(Finding("WARN", 0, f"{title_slides} title slides; a builder copied from an earlier deck may call it twice"))
    return found


def numbered_lines(path):
    rows = []
    for n, slide in enumerate(Presentation(path).slides, 1):
        for shape in _walk(slide.shapes):
            if getattr(shape, "has_table", False) and shape.has_table:
                rows += [(n, "table", " | ".join(c.text.replace("\n", " ") for c in row.cells)) for row in shape.table.rows]
            elif shape.has_text_frame:
                rows += [(n, "slide", p.text) for p in shape.text_frame.paragraphs if p.text.strip() != str(n)]
        rows += [(n, "notes", line) for line in _notes(slide).splitlines()]
    return [row for row in rows if re.search(r"\d", row[2])]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("deck", type=Path)
    parser.add_argument("--people", default="", help="comma-separated names of people in the room; they never appear in a deck")
    parser.add_argument("--dump-numbers", action="store_true")
    args = parser.parse_args()
    if not args.deck.is_file():
        print(f"no such file: {args.deck}", file=sys.stderr)
        return 2
    if args.dump_numbers:
        print("slide\tlocation\tline")
        for n, location, line in numbered_lines(args.deck):
            print(f"{n}\t{location}\t{line}")
        return 0
    found = check(args.deck, tuple(p for p in args.people.split(",") if p))
    for f in found:
        print(f"{f.level} {'deck' if f.slide == 0 else f'slide {f.slide}'}: {f.message}")
    errors = sum(f.level == "ERROR" for f in found)
    print(f"{errors} errors, {len(found) - errors} warnings")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
