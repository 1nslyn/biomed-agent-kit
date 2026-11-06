"""Every check is proved by a forged deck that breaks it. Run: python3 -m pytest scripts/test_check_deck.py"""

import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))
import check_deck  # noqa: E402
import deck_style as ds  # noqa: E402


def clean_deck():
    prs = ds.new_deck()
    ds.title_slide(prs, "Weekly Meeting November 13th", ["A. Presenter", "Nov 6 – 13, 2026"], "Covers one week.")
    s = ds.page(prs, "Project A: test AUC", "Trained on cohort A, tested on 80 samples: 61.3% to **64.1%**.", "Model B, mean of 5 seeds.")
    ds.table(s, [["Experiment", "Result"], ["Model B, 80 test samples", "**64.1%**"]], 0.5, 1.6, [6, 6])
    ds.box(s, 0.5, 3.0, 3, 1, ["A plain box"])
    ds.arrow(s, 3.5, 3.5, 5, 3.5)
    return prs


def page_two(prs):
    return list(prs.slides)[1]


def add_run(slide, content, name="Times New Roman", rgb=None):
    tf = slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(0.5)).text_frame
    r = tf.paragraphs[0].add_run()
    r.text = content
    if name:
        r.font.name = name
    r.font.size = Pt(14)
    if rgb:
        r.font.color.rgb = RGBColor.from_string(rgb)
    return r


def findings(prs, tmp_path, **kw):
    path = tmp_path / "deck.pptx"
    prs.save(path)
    return check_deck.check(path, **kw)


def has(found, level, pattern):
    return any(f.level == level and re.search(pattern, f.message) for f in found)


def rewrite_part(src, dst, part, edit):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            zout.writestr(item, edit(data) if item.filename == part else data)


def test_clean_deck_has_no_findings(tmp_path):
    assert findings(clean_deck(), tmp_path) == []


@pytest.mark.parametrize("name", ["Arial", None])
def test_wrong_or_inherited_font_warns(tmp_path, name):
    prs = clean_deck()
    add_run(page_two(prs), "Some text", name=name)
    assert has(findings(prs, tmp_path), "WARN", "font")


def test_font_inside_a_table_cell_is_checked(tmp_path):
    prs = clean_deck()
    table = next(s for s in page_two(prs).shapes if s.has_table).table
    table.cell(1, 0).text = "Inherited font"
    assert has(findings(prs, tmp_path), "WARN", "font")


@pytest.mark.parametrize("rgb", ["C00000", "EB6834", "F59E0B"])
def test_red_orange_amber_text_warns(tmp_path, rgb):
    prs = clean_deck()
    add_run(page_two(prs), "Warm text", rgb=rgb)
    assert has(findings(prs, tmp_path), "WARN", "red or orange")


def test_navy_and_gray_text_are_allowed(tmp_path):
    prs = clean_deck()
    add_run(page_two(prs), "Navy", rgb="1F3A5F")
    add_run(page_two(prs), "Gray", rgb="5E5D59")
    assert findings(prs, tmp_path) == []


def test_missing_notes_warns(tmp_path):
    prs = clean_deck()
    prs.slides.add_slide(prs.slide_layouts[6])
    assert has(findings(prs, tmp_path), "WARN", "notes")


@pytest.mark.parametrize("where", ["slide", "notes"])
def test_em_dash_warns(tmp_path, where):
    prs = clean_deck()
    if where == "slide":
        add_run(page_two(prs), "Gain — large")
    else:
        page_two(prs).notes_slide.notes_text_frame.text = "Gain — large"
    assert has(findings(prs, tmp_path), "WARN", "em dash")


def test_en_dash_range_is_allowed(tmp_path):
    prs = clean_deck()
    add_run(page_two(prs), "Range 71–96%")
    assert findings(prs, tmp_path) == []


@pytest.mark.parametrize("where", ["slide", "notes"])
def test_person_in_the_room_is_an_error(tmp_path, where):
    prs = clean_deck()
    if where == "slide":
        add_run(page_two(prs), "Ann's checklist")
    else:
        page_two(prs).notes_slide.notes_text_frame.text = "Send Ann the reply"
    assert has(findings(prs, tmp_path, people=("Ann",)), "ERROR", "Ann")


def test_longer_word_is_not_a_person(tmp_path):
    prs = clean_deck()
    add_run(page_two(prs), "Annual run")
    assert findings(prs, tmp_path, people=("Ann",)) == []


def test_default_banded_table_warns(tmp_path):
    prs = clean_deck()
    page_two(prs).shapes.add_table(2, 2, Inches(1), Inches(5), Inches(4), Inches(1))
    assert has(findings(prs, tmp_path), "WARN", "table")


def test_theme_styled_shape_warns(tmp_path):
    prs = clean_deck()
    page_two(prs).shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(5), Inches(1), Inches(1))
    assert has(findings(prs, tmp_path), "WARN", "shadow")


def test_shape_off_the_slide_is_an_error(tmp_path):
    prs = clean_deck()
    ds.box(page_two(prs), 12.5, 3, 3, 1, ["Too wide"])
    assert has(findings(prs, tmp_path), "ERROR", "off the slide")


def test_duplicate_shape_id_is_an_error(tmp_path):
    src = tmp_path / "src.pptx"
    clean_deck().save(src)
    dst = tmp_path / "dst.pptx"

    def edit(data):
        return re.sub(rb'(<p:cNvPr id=")\d+', rb"\g<1>2", data)

    rewrite_part(src, dst, "ppt/slides/slide2.xml", edit)
    assert has(check_deck.check(dst), "ERROR", "shape id")


def test_notes_master_list_after_slide_list_is_an_error(tmp_path):
    src = tmp_path / "src.pptx"
    clean_deck().save(src)
    dst = tmp_path / "dst.pptx"

    def edit(data):
        return data.replace(b"</p:sldIdLst>", b'</p:sldIdLst><p:notesMasterIdLst><p:notesMasterId r:id="rId99"/></p:notesMasterIdLst>')

    rewrite_part(src, dst, "ppt/presentation.xml", edit)
    assert has(check_deck.check(dst), "ERROR", "Repair")


def test_content_type_for_a_missing_part_is_an_error(tmp_path):
    src = tmp_path / "src.pptx"
    clean_deck().save(src)
    dst = tmp_path / "dst.pptx"

    def edit(data):
        return data.replace(b"</Types>", b'<Override PartName="/ppt/slideMasters/slideMaster9.xml" ContentType="application/xml"/></Types>')

    rewrite_part(src, dst, "[Content_Types].xml", edit)
    assert has(check_deck.check(dst), "ERROR", "slideMaster9")


@pytest.mark.parametrize("line", ["Gain +0.034", "AUC 0.81", "AUC of 0.81", "SE ±0.009"])
def test_decimal_metric_is_a_warning(tmp_path, line):
    prs = clean_deck()
    add_run(page_two(prs), line)
    assert has(findings(prs, tmp_path), "WARN", "percent")


@pytest.mark.parametrize("line", ["Gain +3.4%", "Gain +0.13%", "60% AUC", "p = 0.034", "lr 1e-5"])
def test_percent_and_p_value_pass(tmp_path, line):
    prs = clean_deck()
    add_run(page_two(prs), line)
    assert findings(prs, tmp_path) == []


def test_rejected_wording_warns_on_slides_only(tmp_path):
    prs = clean_deck()
    add_run(page_two(prs), "The treatment arm")
    assert has(findings(prs, tmp_path), "WARN", "arm")
    prs = clean_deck()
    page_two(prs).notes_slide.notes_text_frame.text = "The treatment arm"
    assert findings(prs, tmp_path) == []


def test_second_title_slide_warns(tmp_path):
    prs = clean_deck()
    ds.title_slide(prs, "Again", ["A. Presenter"], "Copied builder.")
    assert has(findings(prs, tmp_path), "WARN", "title slide")


def test_dump_numbers_lists_slide_and_notes_lines(tmp_path):
    path = tmp_path / "deck.pptx"
    clean_deck().save(path)
    rows = check_deck.numbered_lines(path)
    assert (2, "slide", "Trained on cohort A, tested on 80 samples: 61.3% to 64.1%.") in rows
    assert (2, "table", "Model B, 80 test samples | 64.1%") in rows
    assert all(re.search(r"\d", row[2]) for row in rows)


def test_cli_exit_codes(tmp_path):
    good = tmp_path / "good.pptx"
    clean_deck().save(good)
    bad_deck = clean_deck()
    ds.box(page_two(bad_deck), 12.5, 3, 3, 1, ["Too wide"])
    bad = tmp_path / "bad.pptx"
    bad_deck.save(bad)
    run = lambda *a: subprocess.run([sys.executable, str(HERE / "check_deck.py"), *a], capture_output=True, text=True)
    assert run(str(good)).returncode == 0
    assert run(str(bad)).returncode == 1
    assert run(str(tmp_path / "absent.pptx")).returncode == 2


def test_shrunken_chart_warns(tmp_path):
    from PIL import Image

    png = tmp_path / "chart.png"
    Image.new("RGB", (2460, 900), "white").save(png)
    prs = clean_deck()
    page_two(prs).shapes.add_picture(str(png), Inches(1), Inches(4.5), width=Inches(4))
    assert has(findings(prs, tmp_path), "WARN", "dpi")


def test_chart_at_slide_size_passes(tmp_path):
    from PIL import Image

    png = tmp_path / "chart.png"
    Image.new("RGB", (2460, 900), "white").save(png)
    prs = clean_deck()
    ds.figure(page_two(prs), str(png), y=3)
    assert findings(prs, tmp_path) == []


def test_style_departures_do_not_fail_the_cli(tmp_path):
    prs = clean_deck()
    add_run(page_two(prs), "Text", name="Arial")
    path = tmp_path / "styled.pptx"
    prs.save(path)
    assert subprocess.run([sys.executable, str(HERE / "check_deck.py"), str(path)], capture_output=True).returncode == 0


def test_font_check_follows_deck_style(tmp_path, monkeypatch):
    monkeypatch.setattr(check_deck, "FONT", "Arial")
    assert has(findings(clean_deck(), tmp_path), "WARN", "font")


def test_example_deck_builds_clean(tmp_path):
    for name in ("deck_style.py", "chart_style.py"):
        shutil.copy(HERE / name, tmp_path / name)
    shutil.copytree(HERE.parent / "examples", tmp_path, dirs_exist_ok=True)
    (tmp_path / "figs/deck").mkdir(parents=True)
    joined = (tmp_path / "chart_style.py").read_text() + (tmp_path / "plot_deck.py").read_text()
    subprocess.run([sys.executable, "-", "figs/deck"], input=joined, text=True, cwd=tmp_path, check=True, capture_output=True)
    subprocess.run([sys.executable, "build_deck.py"], cwd=tmp_path, check=True, capture_output=True)
    assert check_deck.check(tmp_path / "example.pptx") == []


def test_title_too_long_for_one_line_warns(tmp_path):
    prs = clean_deck()
    ds.page(prs, "Project A: " + "a very long title " * 5, "Takeaway.", "Notes.")
    assert has(findings(prs, tmp_path), "WARN", "title")


def test_todos_slide_sets_project_headings_apart(tmp_path):
    prs = clean_deck()
    s = ds.todos(prs, {"Project A": ["Finish the rerun."], "Project B": ["Check the labels."]}, "Notes.")
    assert findings(prs, tmp_path) == []
    paragraphs = [p for shape in s.shapes if shape.has_text_frame for p in shape.text_frame.paragraphs]
    bold = {p.text for p in paragraphs if p.runs and p.runs[0].font.bold}
    assert {"Project A", "Project B"} <= bold
    assert "Finish the rerun." not in bold
