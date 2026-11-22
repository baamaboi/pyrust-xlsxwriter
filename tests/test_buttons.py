"""Form-control buttons."""

import copy
import re
import zipfile
from pathlib import Path

import pytest

import pyrust_xlsxwriter as px

CORPUS = Path(__file__).parent / "corpus"


def normalize_vml(xml: str) -> str:
    """Ignore whitespace between/inside tags and quote style (as Excel varies them)."""
    xml = xml.replace("'", '"')
    xml = re.sub(r">\s+<", "><", xml)
    xml = re.sub(r";\s+", ";", xml)
    return re.sub(r"\s+", " ", xml).replace("> ", ">").strip()


def button_shapes(xml: str) -> list[str]:
    return re.findall(r"<v:shape .*?</v:shape>", xml, re.S)


def test_button_defaults_and_kwargs():
    button = px.Button()
    assert button.caption is None and button.macro_name is None
    assert (button.width, button.height) == (64, 20)
    assert button.alt_text == ""
    assert button.object_movement == px.ObjectMovement.MoveAndSizeWithCells
    button = px.Button(
        caption="Run",
        macro_name="say_hello",
        width=100,
        height=30,
        alt_text="Runs it",
        object_movement=px.ObjectMovement.DontMoveOrSizeWithCells,
    )
    assert (button.caption, button.macro_name, button.width, button.height) == ("Run", "say_hello", 100, 30)
    assert button.alt_text == "Runs it"
    assert button.object_movement == px.ObjectMovement.DontMoveOrSizeWithCells
    button.caption = None
    assert button.caption is None
    with pytest.raises(TypeError, match="unexpected keyword argument 'macro'"):
        px.Button(macro="x")  # pyrefly: ignore[unexpected-keyword]


def test_button_validation():
    button = px.Button()
    with pytest.raises(px.ParameterError, match="255"):
        button.caption = "x" * 256
    with pytest.raises(px.ParameterError, match="255"):
        button.alt_text = "x" * 256
    for name in ("width", "height"):
        with pytest.raises(px.ParameterError):
            setattr(button, name, 0)


def test_matches_excel_reference(out, reopen, part):
    """Same VML as Excel's button07.xlsm (a button with a macro at C2)."""
    wb = px.Workbook()
    wb.add_vba_project(CORPUS / "vbaProject.bin")
    ws = wb.add_worksheet()
    ws.insert_button(1, 2, px.Button(caption="Hello", macro_name="say_hello"))
    path = out("button.xlsm")
    wb.save(path)
    reopen(path)
    with zipfile.ZipFile(path) as zf:
        assert "xl/vbaProject.bin" in zf.namelist()
    with zipfile.ZipFile(CORPUS / "button07.xlsm") as zf:
        expected = zf.read("xl/drawings/vmlDrawing1.vml").decode()
    assert normalize_vml(part(path, "xl/drawings/vmlDrawing1.vml")) == normalize_vml(expected)


def test_insert_buttons(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_button("B2")  # default button
    ws.insert_button(4, 1, px.Button(caption="Offset", width=80, height=30), x_offset=10, y_offset=5)
    ws.insert_button("B8", px.Button(alt_text="Alternative"), 3)
    ws.insert_button_with_offset("B10", px.Button(caption="With offset"), 2, 4)
    ws.insert_button_with_offset(12, "B", button=px.Button(macro_name="m"), x_offset=0, y_offset=0)
    path = out()
    wb.save(path)
    reopen(path)
    shapes = button_shapes(part(path, "xl/drawings/vmlDrawing1.vml"))
    assert len(shapes) == 5
    assert ">Button 1<" in shapes[0]
    assert "width:60pt;height:22.5pt" in shapes[1] and ">Offset<" in shapes[1]
    assert re.search(r"<x:Anchor>\s*1, 10, 4, 5,", shapes[1])
    assert 'alt="Alternative"' in shapes[2]
    assert re.search(r"<x:Anchor>\s*1, 2, 9, 4,", shapes[3])
    assert "<x:FmlaMacro>[0]!m</x:FmlaMacro>" in shapes[4]


def test_insert_button_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError, match="missing required argument: 'x_offset'"):
        ws.insert_button_with_offset("A1", px.Button())  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError):
        ws.insert_button("A1", "Run")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError):
        ws.insert_button("A1", px.Button(), x_offset=-1)
    with pytest.raises(px.RangeError):
        ws.insert_button(1_048_576, 0)


def test_button_copy_eq_repr():
    button = px.Button(caption="Run", macro_name="go")
    dup = button.copy()
    assert dup == button and dup is not button
    dup.caption = "Stop"
    assert dup != button and button.caption == "Run"
    assert copy.copy(button) == button and copy.deepcopy(button) == button
    assert button != "Run"
    assert repr(button) == "Button(caption='Run', macro_name='go')"
