from pathlib import Path

from PIL import Image, ImageDraw

from gns_app.services.visual_classifier import PageVisualAnalyzer


def _save_page(path: Path, *, table: bool) -> None:
    image = Image.new("L", (800, 1100), "white")
    draw = ImageDraw.Draw(image)
    for y in range(100, 950, 28):
        draw.line((90, y, 700, y), fill="black", width=2)
    if table:
        for x in range(90, 701, 48):
            draw.line((x, 250, x, 940), fill="black", width=2)
    image.save(path)


def test_detects_dense_sti_table_layout(tmp_path: Path):
    path = tmp_path / "decision.png"
    _save_page(path, table=True)

    result = PageVisualAnalyzer().analyze(path)

    assert result.decision_layout
    assert result.confidence >= 0.86
    assert result.horizontal_line_groups >= 10
    assert result.vertical_line_groups >= 8


def test_text_lines_without_table_are_not_decision_layout(tmp_path: Path):
    path = tmp_path / "letter.png"
    _save_page(path, table=False)

    result = PageVisualAnalyzer().analyze(path)

    assert not result.decision_layout
