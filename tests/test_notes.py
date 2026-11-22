"""Cell notes: the Note class and insert_note / show_all_notes /
set_default_note_author."""

import copy
import re

import pytest

import pyrust_xlsxwriter as px


def vml_shapes(xml: str) -> list[str]:
    return re.findall(r"<v:shape .*?</v:shape>", xml, re.S)


def test_note_defaults():
    note = px.Note("Hello")
    assert note.text == "Hello"
    assert note.author is None
    assert note.add_author_prefix is True
    assert (note.width, note.height) == (128, 74)
    assert note.visible is None
    assert note.background_color is None
    assert note.font_name is None and note.font_size is None and note.font_family is None
    assert note.format is None
    assert note.alt_text == ""
    assert note.object_movement == px.ObjectMovement.DontMoveOrSizeWithCells
    assert px.Note().text == ""


def test_note_kwargs_and_setters():
    fmt = px.Format(italic=True)
    note = px.Note(
        "Hi",
        author="Ana",
        add_author_prefix=False,
        width=200,
        height=100,
        visible=True,
        background_color="#FF0000",
        font_name="Arial",
        font_size=12,
        font_family=2,
        format=fmt,
        alt_text="alt",
        object_movement=px.ObjectMovement.MoveAndSizeWithCells,
    )
    assert note.author == "Ana" and note.add_author_prefix is False
    assert (note.width, note.height) == (200, 100)
    assert note.visible is True
    assert note.background_color == px.Color("#FF0000")
    assert (note.font_name, note.font_size, note.font_family) == ("Arial", 12.0, 2)
    assert note.format is fmt  # live: the same object
    assert note.alt_text == "alt"
    assert note.object_movement == px.ObjectMovement.MoveAndSizeWithCells
    note.text = "changed"
    note.reset_text("reset")
    assert note.text == "reset"
    note.author = None
    note.visible = None
    note.format = None
    assert (note.author, note.visible, note.format) == (None, None, None)
    with pytest.raises(TypeError, match="unexpected keyword argument 'colour'"):
        px.Note("x", colour="red")  # pyrefly: ignore[unexpected-keyword]


def test_note_validation():
    note = px.Note("x")
    with pytest.raises(px.ParameterError, match="52"):
        note.author = "a" * 53
    note.author = "a" * 52
    for name in ("width", "height"):
        with pytest.raises(px.ParameterError):
            setattr(note, name, 0)
        with pytest.raises(px.ParameterError):
            setattr(note, name, 2_000_000)
    with pytest.raises(px.ParameterError):
        note.font_size = 0
    with pytest.raises(px.ParameterError, match="255"):
        note.alt_text = "x" * 256
    with pytest.raises(px.ParameterError):
        note.background_color = "not-a-colour"
    with pytest.raises(TypeError):
        note.format = "bold"  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.ParameterError):
        px.Note("x", author="b" * 60)


def test_insert_note(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write("B2", "value")
    ws.insert_note("B2", px.Note("Check this", author="Ana"))
    ws.insert_note(2, 2, "plain text note")
    ws.insert_note(3, "D", note=px.Note("no prefix", author="Bob", add_author_prefix=False))
    path = out()
    wb.save(path)
    sheet = reopen(path).active
    assert sheet["B2"].value == "value"
    assert sheet["B2"].comment.text == "Ana:\nCheck this"
    assert sheet["B2"].comment.author == "Ana"
    assert sheet["C3"].comment.text == "Author:\nplain text note"
    assert sheet["D4"].comment.text == "no prefix"
    assert sheet["D4"].comment.author == "Bob"
    comments = part(path, "xl/comments1.xml")
    assert "<authors><author>Author</author><author>Ana</author><author>Bob</author></authors>" in comments


def test_default_author_and_show_all(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_default_note_author("Default Person")
    ws.show_all_notes()
    ws.insert_note("A1", "one")
    ws.insert_note("A3", px.Note("hidden", visible=False))
    path = out()
    wb.save(path)
    sheet = reopen(path).active
    assert sheet["A1"].comment.author == "Default Person"
    assert sheet["A1"].comment.text == "Default Person:\none"
    shapes = vml_shapes(part(path, "xl/drawings/vmlDrawing1.vml"))
    assert "visibility:visible" in shapes[0]
    assert "visibility:hidden" in shapes[1]
    with pytest.raises(px.ParameterError, match="52"):
        ws.set_default_note_author("x" * 53)


def test_show_all_notes_off(out, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.show_all_notes(True)
    ws.show_all_notes(enable=False)
    ws.insert_note("A1", "one")
    ws.insert_note("A3", px.Note("shown", visible=True))
    path = out()
    wb.save(path)
    shapes = vml_shapes(part(path, "xl/drawings/vmlDrawing1.vml"))
    assert "visibility:hidden" in shapes[0]
    assert "visibility:visible" in shapes[1]


def test_note_appearance(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_note(
        "B2",
        px.Note(
            "styled",
            author="A",
            width=200,
            height=100,
            background_color="red",
            font_name="Arial",
            font_size=12,
            alt_text="Alternative",
        ),
    )
    ws.insert_note("B5", px.Note("default"))
    path = out()
    wb.save(path)
    reopen(path)
    vml = part(path, "xl/drawings/vmlDrawing1.vml")
    styled, default = vml_shapes(vml)
    assert 'fillcolor="#ff0000"' in styled and "width:150pt;height:75pt" in styled
    assert "Alternative" in styled
    assert 'fillcolor="#ffffe1"' in default and "width:96pt;height:55.5pt" in default
    comments = part(path, "xl/comments1.xml")
    assert '<sz val="12"/>' in comments and '<rFont val="Arial"/>' in comments
    assert '<sz val="8"/>' in comments and '<rFont val="Tahoma"/>' in comments


def test_note_format_is_live(out, part):
    fmt = px.Format(font_name="Courier New", font_size=9)
    note = px.Note("fmt", format=fmt)
    fmt.font_size = 15  # changed after assignment: still used
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_note("A1", note)
    path = out()
    wb.save(path)
    comments = part(path, "xl/comments1.xml")
    assert '<rFont val="Courier New"/>' in comments and '<sz val="15"/>' in comments


def test_insert_note_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(TypeError, match="Note or a str"):
        ws.insert_note("A1", 5)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="missing required argument"):
        ws.insert_note("A1")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError):
        ws.insert_note("A1", "x" * 32_767)
    with pytest.raises(px.RangeError):
        ws.insert_note(0, 16_384, "x")
    # Notes at the last row/column are allowed.
    ws.insert_note(1_048_575, 16_383, "corner")


def test_note_copy_eq_repr():
    note = px.Note("a", author="Ana", visible=True, format=px.Format(bold=True))
    dup = note.copy()
    assert dup == note and dup is not note
    assert dup.format is not note.format and dup.format == note.format  # deep copy
    shallow = copy.copy(note)
    assert shallow.format is note.format
    assert copy.deepcopy(note).format is not note.format
    assert dup.format is not None
    dup.format.bold = False
    assert dup != note
    assert px.Note("a") != px.Note("b")
    assert px.Note("a") != "a"
    assert repr(note) == "Note('a', author='Ana', visible=True)"
