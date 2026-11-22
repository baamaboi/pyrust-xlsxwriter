"""Shapes (text boxes) and their formatting classes."""

import copy
import re

import pytest

import pyrust_xlsxwriter as px

EMU = 9525


def shapes_xml(path, part) -> list[str]:
    return re.findall(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", part(path, "xl/drawings/drawing1.xml"))


def save_one(shape, out, reopen, part, cell="B2"):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_shape(cell, shape)
    path = out()
    wb.save(path)
    reopen(path)
    return shapes_xml(path, part)[0]


# ---------------------------------------------------------------------------
# Enums
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    ("enum", "members"),
    [
        (
            px.ShapeLineDashType,
            ["Solid", "RoundDot", "SquareDot", "Dash", "DashDot", "LongDash", "LongDashDot", "LongDashDotDot"],
        ),
        (px.ShapeGradientFillType, ["Linear", "Radial", "Rectangular", "Path"]),
        (px.ShapeTextDirection, ["Horizontal", "Rotate270", "Rotate90", "Rotate90EastAsian", "Stacked"]),
        (px.ShapeTextHorizontalAlignment, ["Default", "Left", "Center", "Right"]),
        (px.ShapeTextVerticalAlignment, ["Top", "Middle", "Bottom", "TopCentered", "MiddleCentered", "BottomCentered"]),
    ],
)
def test_enums(enum, members):
    for i, name in enumerate(members):
        assert int(getattr(enum, name)) == i


def test_pattern_enum_complete():
    names = [n for n in dir(px.ShapePatternFillType) if not n.startswith("_")]
    assert len(names) == 48
    assert int(px.ShapePatternFillType.SolidDiamondGrid) == 47


# ---------------------------------------------------------------------------
# Leaf classes
# ---------------------------------------------------------------------------


def test_shape_line():
    line = px.ShapeLine()
    assert (line.color, line.width, line.dash_type, line.transparency, line.hidden) == (
        None,
        0.75,
        px.ShapeLineDashType.Solid,
        0,
        False,
    )
    line = px.ShapeLine(color="red", width=2.5, dash_type=px.ShapeLineDashType.Dash, transparency=40, hidden=True)
    assert line.color == px.Color.Red and line.width == 2.5 and line.transparency == 40
    with pytest.raises(px.ParameterError):
        line.width = 1585
    with pytest.raises(px.ParameterError):
        line.width = -1
    with pytest.raises(px.ParameterError):
        line.transparency = 101
    with pytest.raises(px.ParameterError):
        line.color = "nope"
    assert repr(px.ShapeLine(width=2.0)) == "ShapeLine(width=2.0)"


def test_fills_font_text():
    solid = px.ShapeSolidFill(color=(255, 0, 0), transparency=50)
    assert solid.color == px.Color("#FF0000") and solid.transparency == 50
    with pytest.raises(px.ParameterError):
        solid.transparency = -1
    pattern = px.ShapePatternFill()
    assert pattern.pattern == px.ShapePatternFillType.Dotted5Percent
    assert pattern.background_color is None and pattern.foreground_color is None
    pattern.foreground_color = px.Color.theme(4, 0)
    assert pattern.foreground_color == px.Color.theme(4, 0)
    font = px.ShapeFont(
        bold=True,
        italic=True,
        underline=True,
        strikethrough=True,
        name="Arial",
        size=14,
        color="blue",
        right_to_left=False,
        pitch_family=34,
        character_set=0,
    )
    assert font.bold and font.italic and font.underline and font.strikethrough
    assert (font.name, font.size, font.color) == ("Arial", 14.0, px.Color.Blue)
    with pytest.raises(px.ParameterError):
        font.size = 0
    text = px.ShapeText(horizontal_alignment=px.ShapeTextHorizontalAlignment.Center)
    assert text.vertical_alignment == px.ShapeTextVerticalAlignment.Top
    assert text.direction == px.ShapeTextDirection.Horizontal
    assert repr(text) == "ShapeText(horizontal_alignment=ShapeTextHorizontalAlignment.Center)"


def test_gradient_fill_and_stops():
    fill = px.ShapeGradientFill()
    assert fill.type == px.ShapeGradientFillType.Linear
    assert fill.angle == 90 and fill.gradient_stops == []
    stop = px.ShapeGradientStop("#DDEBCF", 0)
    fill = px.ShapeGradientFill(
        type=px.ShapeGradientFillType.Radial,
        gradient_stops=[stop, ("#9CB86E", 50), ("#156B13", 100)],
        angle=45,
    )
    assert fill.type == px.ShapeGradientFillType.Radial and fill.angle == 45
    stops = fill.gradient_stops
    assert stops[0] is stop  # live
    assert [s.position for s in stops] == [0, 50, 100]
    assert stops[1].color == px.Color("#9CB86E")
    fill.type = px.ShapeGradientFillType.Path
    assert fill.type == px.ShapeGradientFillType.Path
    with pytest.raises(px.ParameterError, match="2 to 10"):
        fill.gradient_stops = [stop]
    with pytest.raises(px.ParameterError, match="2 to 10"):
        fill.gradient_stops = [stop] * 11
    with pytest.raises(px.ParameterError):
        fill.gradient_stops = [("red", 0), ("blue", 101)]
    with pytest.raises(TypeError):
        fill.gradient_stops = ["red", "blue"]  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.ParameterError):
        fill.angle = 360
    with pytest.raises(px.ParameterError):
        px.ShapeGradientStop("red", -1)
    with pytest.raises(px.ParameterError):
        stop.position = 101
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.ShapeGradientFill(kind=px.ShapeGradientFillType.Linear)  # pyrefly: ignore[unexpected-keyword]
    fill.gradient_stops = []
    assert fill.gradient_stops == []


# ---------------------------------------------------------------------------
# ShapeFormat and Shape
# ---------------------------------------------------------------------------


def test_shape_format_fills_are_exclusive():
    fmt = px.ShapeFormat()
    assert isinstance(fmt.line, px.ShapeLine)
    assert (fmt.no_line, fmt.no_fill, fmt.solid_fill, fmt.pattern_fill, fmt.gradient_fill) == (
        False,
        False,
        None,
        None,
        None,
    )
    fmt.solid_fill = px.ShapeSolidFill(color="red")
    fmt.pattern_fill = px.ShapePatternFill()
    assert fmt.solid_fill is None and fmt.pattern_fill is not None
    fmt.gradient_fill = px.ShapeGradientFill(gradient_stops=[("red", 0), ("blue", 100)])
    assert fmt.pattern_fill is None
    fmt.no_fill = True
    assert fmt.gradient_fill is None and fmt.no_fill
    fmt.solid_fill = px.ShapeSolidFill()
    assert not fmt.no_fill
    fmt.no_line = True
    fmt.line = px.ShapeLine(color="red")
    assert not fmt.no_line
    fmt.line = None  # resets to a default line
    assert fmt.line == px.ShapeLine()


def test_shape_defaults_and_kwargs():
    shape = px.Shape()
    assert shape.text == "" and shape.text_link is None
    assert (shape.width, shape.height) == (192, 120)
    assert isinstance(shape.format, px.ShapeFormat)
    assert isinstance(shape.font, px.ShapeFont) and isinstance(shape.text_options, px.ShapeText)
    assert shape.url is None and shape.alt_text == ""
    assert shape.object_movement == px.ObjectMovement.MoveAndSizeWithCells
    font = px.ShapeFont(bold=True)
    shape = px.Shape.textbox(
        text="Hi",
        text_link="=A1",
        width=300,
        height=50,
        format=px.ShapeLine(color="red"),
        font=font,
        text_options=px.ShapeText(direction=px.ShapeTextDirection.Stacked),
        url="https://example.com",
        alt_text="box",
        object_movement=px.ObjectMovement.DontMoveOrSizeWithCells,
    )
    assert shape.text == "Hi" and shape.text_link == "=A1"
    assert shape.format.line.color == px.Color.Red  # a ShapeLine is wrapped in a ShapeFormat
    assert shape.font is font
    assert shape.url is not None and shape.url.link == "https://example.com"
    shape.text_link = px.Formula("=Sheet1!B2")
    assert shape.text_link == "=Sheet1!B2"
    shape.format = px.ShapeSolidFill(color="blue")
    assert shape.format.solid_fill is not None and shape.format.solid_fill.color == px.Color.Blue
    shape.format = None
    shape.font = None
    shape.text_options = None
    assert shape.format == px.ShapeFormat() and shape.font == px.ShapeFont()
    assert shape.text_options == px.ShapeText()


def test_shape_validation():
    shape = px.Shape()
    with pytest.raises(px.ParameterError):
        shape.width = 4
    with pytest.raises(px.ParameterError):
        shape.height = 2_000_000
    with pytest.raises(px.ParameterError, match="255"):
        shape.alt_text = "x" * 256
    with pytest.raises(px.ParameterError):
        shape.url = "nowhere"
    with pytest.raises(px.ParameterError, match="empty"):
        shape.text_link = "="
    with pytest.raises(TypeError):
        shape.text_link = 5  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="ShapeFormat"):
        shape.format = "red"  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        shape.font = px.ShapeText()  # pyrefly: ignore[bad-argument-type]


def test_nested_objects_are_live(out, reopen, part):
    shape = px.Shape.textbox(text="Live")
    shape.format.line.color = "red"
    shape.format.line.dash_type = px.ShapeLineDashType.LongDash
    shape.format.line.width = 2.25
    shape.format.solid_fill = px.ShapeSolidFill(color="#00FF00")
    assert shape.format.solid_fill is not None
    shape.format.solid_fill.transparency = 25
    shape.font.bold = True
    shape.font.size = 14
    shape.font.name = "Arial"
    shape.font.color = "blue"
    shape.text_options.vertical_alignment = px.ShapeTextVerticalAlignment.Middle
    shape.text_options.horizontal_alignment = px.ShapeTextHorizontalAlignment.Center
    shape.text_options.direction = px.ShapeTextDirection.Rotate90
    xml = save_one(shape, out, reopen, part)
    match = re.search(r"<a:ln .*?</a:ln>", xml)
    assert match
    line = match.group(0)
    assert 'w="28575"' in line and '<a:srgbClr val="FF0000"/>' in line
    assert '<a:prstDash val="lgDash"/>' in line
    assert '<a:solidFill><a:srgbClr val="00FF00"><a:alpha val="75000"/></a:srgbClr></a:solidFill>' in xml
    assert 'sz="1400" b="1"' in xml and '<a:latin typeface="Arial"/>' in xml
    assert '<a:srgbClr val="0000FF"/>' in xml
    assert 'vert="vert"' in xml and 'anchor="ctr"' in xml and 'algn="ctr"' in xml
    assert "<a:t>Live</a:t>" in xml


def test_shape_formats_in_xml(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    pattern = px.Shape(
        format=px.ShapePatternFill(
            pattern=px.ShapePatternFillType.Plaid, foreground_color="red", background_color="yellow"
        )
    )
    gradient = px.Shape(format=px.ShapeGradientFill(gradient_stops=[("red", 0), ("blue", 100)], angle=45))
    radial = px.Shape(
        format=px.ShapeGradientFill(type=px.ShapeGradientFillType.Radial, gradient_stops=[("red", 0), ("blue", 100)])
    )
    bare = px.Shape(format=px.ShapeFormat(no_line=True, no_fill=True))
    hidden_line = px.Shape(format=px.ShapeLine(hidden=True))
    for row, shape in enumerate([pattern, gradient, radial, bare, hidden_line]):
        ws.insert_shape(row * 10, 0, shape)
    path = out()
    wb.save(path)
    reopen(path)
    items = shapes_xml(path, part)
    assert (
        '<a:pattFill prst="plaid"><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr><a:bgClr><a:srgbClr val="FFFF00"/>'
        in items[0]
    )
    assert '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000">' in items[1]
    assert '<a:lin ang="2700000" scaled="0"/>' in items[1]
    assert '<a:path path="circle">' in items[2]
    assert '</a:prstGeom><a:noFill/><a:ln w="9525" cmpd="sng"><a:noFill/></a:ln>' in items[3]
    assert re.search(r"<a:ln [^>]*><a:noFill/></a:ln>", items[4])
    assert "</a:prstGeom><a:solidFill>" in items[4]  # default fill kept


def test_text_link_url_alt_text(out, reopen, part):
    shape = px.Shape(text="linked", text_link="=Sheet1!$A$1", url="https://example.com/x", alt_text="Description")
    xml = save_one(shape, out, reopen, part)
    assert 'textlink="Sheet1!$A$1"' in xml
    assert 'descr="Description"' in xml and "a:hlinkClick" in xml


def test_insert_shape_positions(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    shape = px.Shape(width=128, height=40)
    ws.insert_shape("B2", shape)
    ws.insert_shape(5, 1, shape, x_offset=10, y_offset=4)
    ws.insert_shape(10, "B", shape, 7)
    ws.insert_shape_with_offset("B20", shape, 3, 2)
    ws.insert_shape_with_offset(25, 1, shape=shape, x_offset=1, y_offset=1)
    path = out()
    wb.save(path)
    reopen(path)
    items = shapes_xml(path, part)
    assert len(items) == 5
    assert "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>" in items[0]
    assert "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row>" in items[0]
    assert f'<a:ext cx="{128 * EMU}" cy="{40 * EMU}"/>' in items[0]
    assert f"<xdr:colOff>{10 * EMU}</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>{4 * EMU}</xdr:rowOff>" in items[1]
    assert f"<xdr:colOff>{7 * EMU}</xdr:colOff><xdr:row>10</xdr:row>" in items[2]
    with pytest.raises(TypeError, match="missing required argument: 'x_offset'"):
        ws.insert_shape_with_offset("A1", shape)  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError):
        ws.insert_shape("A1", shape, x_offset=-5)
    with pytest.raises(TypeError):
        ws.insert_shape("A1", px.ShapeFormat())  # pyrefly: ignore[bad-argument-type]


def test_shape_changes_after_insert(out, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    shape = px.Shape(text="first")
    ws.insert_shape("A1", shape)
    shape.text = "second"
    ws.insert_shape("A20", shape)
    path = out()
    wb.save(path)
    items = shapes_xml(path, part)
    assert "<a:t>first</a:t>" in items[0] and "<a:t>second</a:t>" in items[1]


def test_gradient_without_stops_rejected_on_insert():
    ws = px.Workbook().add_worksheet()
    shape = px.Shape(format=px.ShapeGradientFill())
    with pytest.raises(px.ParameterError, match="2 to 10"):
        ws.insert_shape("A1", shape)


def test_copy_is_deep_and_eq():
    shape = px.Shape(text="a", format=px.ShapeGradientFill(gradient_stops=[("red", 0), ("blue", 100)]))
    dup = shape.copy()
    assert dup == shape
    assert dup.format is not shape.format
    copied, original = dup.format.gradient_fill, shape.format.gradient_fill
    assert copied is not None and original is not None
    assert copied.gradient_stops[0] is not original.gradient_stops[0]
    copied.gradient_stops[0].position = 10
    assert dup != shape
    assert copy.copy(shape) == shape and copy.deepcopy(shape) == shape
    assert px.ShapeLine(color="red") == px.ShapeLine(color=px.Color.Red)
    assert px.ShapeLine(color="red") != px.ShapeLine(color="#FF0000")  # named vs RGB
    assert px.ShapeLine() != px.ShapeSolidFill()
    assert px.ShapeGradientStop("red", 1) != px.ShapeGradientStop("red", 2)
    fmt = px.ShapeFormat(solid_fill=px.ShapeSolidFill(color="red"))
    assert fmt.copy() == fmt and fmt.copy().solid_fill is not fmt.solid_fill
    for obj in (
        px.ShapeLine(),
        px.ShapeSolidFill(),
        px.ShapePatternFill(),
        px.ShapeFont(),
        px.ShapeText(),
        px.ShapeGradientStop("red", 5),
        px.ShapeGradientFill(),
    ):
        assert obj.copy() == obj and copy.deepcopy(obj) == obj
    assert repr(px.Shape(text="t")) == "Shape.textbox(text='t', width=192, height=120)"
    assert "gradient_stops=[ShapeGradientStop(Color.Red, 0)" in repr(shape.format.gradient_fill)
