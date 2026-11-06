"""The one visual style for lab meeting decks: Times New Roman on white, one takeaway under each title, rule-only tables, plain boxes.
Copy this file into the deck's report folder. `**x**` in any text, table cell or box marks a result number: bold and navy.
"""

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

W, H, M = 13.333, 7.5, 0.5
BODY_W = W - 2 * M
FONT = "Times New Roman"
INK, BODY, MUTED, RULE, TINT = (RGBColor.from_string(c) for c in ("0B0B0B", "30302E", "5E5D59", "8A8984", "F3F2EE"))
TITLE_Y, TAKEAWAY_Y, CONTENT_Y, CAPTION_Y = 0.28, 0.86, 1.6, 6.3
RESULT = RGBColor.from_string("1F3A5F")
NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
ALIGN = dict(l=PP_ALIGN.LEFT, c=PP_ALIGN.CENTER, r=PP_ALIGN.RIGHT)


def new_deck():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
    return prs


def _runs(p, line, size, color, bold):
    for i, part in enumerate(str(line).split("**")):
        if part:
            r = p.add_run()
            r.text, r.font.size, r.font.name = part, Pt(size), FONT
            r.font.bold = bold or i % 2 == 1
            r.font.color.rgb = RESULT if i % 2 else color


def fill(tf, lines, size, color=BODY, bold=False, gap=4):
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(gap)
        _runs(p, line, size, color, bold)


def text(slide, lines, x, y, w, h, size=16, color=BODY, bold=False):
    tf = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)).text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    fill(tf, lines, size, color, bold)


def place(shape, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = Inches(x), Inches(y), Inches(w), Inches(h)


def title_slide(prs, title, lines, notes):
    s = prs.slides.add_slide(prs.slide_layouts[0])
    t, sub = s.shapes.title, s.placeholders[1]
    place(t, M, 2.45, BODY_W, 0.9)
    place(sub, M, 3.45, BODY_W, 1.1)
    fill(t.text_frame, [title], 38, INK, True)
    fill(sub.text_frame, lines, 20)
    for tf in (t.text_frame, sub.text_frame):
        tf.margin_left = 0
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
    s.notes_slide.notes_text_frame.text = notes
    return s


def page(prs, title, takeaway, notes):
    s = prs.slides.add_slide(prs.slide_layouts[5])
    t = s.shapes.title
    place(t, M, TITLE_Y, BODY_W, 0.55)
    tf = t.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    fill(tf, [title], 28, INK, True, gap=0)
    tf.paragraphs[0].alignment = PP_ALIGN.LEFT
    if takeaway:
        text(s, [takeaway], M, TAKEAWAY_Y, BODY_W, 0.65, size=18)
    text(s, [str(len(prs.slides))], W - 1.0, H - 0.42, 0.5, 0.25, size=11, color=MUTED)
    s.shapes[-1].text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    s.notes_slide.notes_text_frame.text = notes
    return s


def todos(prs, by_project, notes):
    """The closing slide: each project name in bold, its items under it."""
    s = page(prs, "Todos", "", notes)
    tf = s.shapes.add_textbox(Inches(M), Inches(1.05), Inches(BODY_W), Inches(5.5)).text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for project, items in by_project.items():
        for line, bold in [(project, True), *((item, False) for item in items)]:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.space_after = Pt(6)
            p.space_before = Pt(14 if bold and p is not tf.paragraphs[0] else 0)
            _runs(p, line, 18, INK if bold else BODY, bold)
    return s


def caption(slide, lines, y=CAPTION_Y, x=M, w=BODY_W):
    text(slide, lines, x, y, w, H - y - 0.45, size=14, color=MUTED)


def fit(slide, path, x, y, w, h, alt=None):
    iw, ih = Image.open(path).size
    k = min(w / iw, h / ih)
    pic = slide.shapes.add_picture(path, Inches(x + (w - iw * k) / 2), Inches(y + (h - ih * k) / 2), Inches(iw * k), Inches(ih * k))
    pic._element.nvPicPr.cNvPr.set("descr", alt or slide.shapes.title.text_frame.text)


def figure(slide, path, y=CONTENT_Y, h=None):
    """A chart drawn at 12.3 in wide goes in at its own size, so its fonts keep their point size."""
    iw, ih = Image.open(path).size
    fit(slide, path, M, y, BODY_W, h or BODY_W * ih / iw)


def _rule(tc_pr, tag, width):
    ln = etree.SubElement(tc_pr, qn(f"a:{tag}"), w=str(int(width * 12700)), cap="flat", cmpd="sng", algn="ctr")
    if width:
        etree.SubElement(etree.SubElement(ln, qn("a:solidFill")), qn("a:srgbClr"), val=str(INK))
    else:
        etree.SubElement(ln, qn("a:noFill"))


def _cell(cell, value, align, size, bold, top, bottom, tint):
    cell.margin_top = cell.margin_bottom = Inches(0.035)
    cell.margin_left = cell.margin_right = Inches(0.07)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    for i, line in enumerate(str(value).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN[align]
        _runs(p, line, size, INK if bold else BODY, bold)
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag, width in (("lnL", 0), ("lnR", 0), ("lnT", top), ("lnB", bottom)):
        _rule(tc_pr, tag, width)
    if tint:
        etree.SubElement(etree.SubElement(tc_pr, qn("a:solidFill")), qn("a:srgbClr"), val=str(TINT))


def table(slide, rows, x, y, widths, size=15, row_h=0.36, align=None, strong=()):
    """Rules above and below the header and at the end, nothing else. `strong` rows are bold on a light tint."""
    align = align or "l" * len(widths)
    shape = slide.shapes.add_table(len(rows), len(widths), Inches(x), Inches(y), Inches(sum(widths)), Inches(row_h * len(rows)))
    t = shape.table
    props = t._tbl.tblPr
    for flag in ("firstRow", "bandRow"):
        props.set(flag, "0")
    props.find(qn("a:tableStyleId")).text = NO_STYLE
    for j, w in enumerate(widths):
        t.columns[j].width = Inches(w)
    last = len(rows) - 1
    for i, row in enumerate(rows):
        t.rows[i].height = Inches(row_h)
        for j, value in enumerate(row):
            _cell(t.cell(i, j), value, align[j], size, i == 0 or i in strong, 1.25 if i == 0 else 0, 0.75 if i == 0 else (1.25 if i == last else 0), i in strong)
    return t


def _drop_theme_style(shape):
    """Without the theme's style block a shape has no shadow; every color it uses is set directly."""
    shape._element.remove(shape._element.find(qn("p:style")))


def band(slide, x, y, w, h):
    """A lane behind a row of boxes: tint only, no outline."""
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    _drop_theme_style(b)
    b.fill.solid()
    b.fill.fore_color.rgb = TINT
    b.line.fill.background()


def node(slide, x, y, w, h, head, detail, dashed=False):
    """A pipeline step on a lane: white box, bold name, one short detail."""
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    _drop_theme_style(b)
    b.fill.solid()
    b.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    b.line.color.rgb = RULE
    b.line.width = Pt(1)
    if dashed:
        b.line.dash_style = MSO_LINE.DASH
    tf = b.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.07)
    tf.margin_top = tf.margin_bottom = 0
    if head:
        fill(tf, [head], 15, INK, True, gap=2)
    start = len(tf.paragraphs) if head else 0
    for i, line in enumerate([detail] if detail else []):
        p = tf.add_paragraph() if start else tf.paragraphs[0]
        _runs(p, line, 13, BODY, False)
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER


def box(slide, x, y, w, h, lines, size=15, strong=False):
    b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    _drop_theme_style(b)
    if strong:
        b.fill.solid()
        b.fill.fore_color.rgb = TINT
    else:
        b.fill.background()
    b.line.color.rgb = RULE
    b.line.width = Pt(1)
    tf = b.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.06)
    tf.margin_top = tf.margin_bottom = 0
    fill(tf, lines, size, BODY, False, gap=0)
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
    return b


def arrow(slide, x0, y0, x1, y1, head=True, dashed=False):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    _drop_theme_style(c)
    c.line.color.rgb = RULE
    c.line.width = Pt(1)
    if dashed:
        c.line.dash_style = MSO_LINE.DASH
    if head:
        etree.SubElement(c.line._get_or_add_ln(), qn("a:tailEnd"), type="triangle")
