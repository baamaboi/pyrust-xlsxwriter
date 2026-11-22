"""Conditional formats: every ConditionalFormat* type, rules, thresholds and options."""

import datetime as dt

import pytest

import pyrust_xlsxwriter as px

pytestmark = pytest.mark.filterwarnings("ignore:.*extension is not supported:UserWarning")

T = px.ConditionalFormatType
CellRule = px.ConditionalFormatCellRule
RED = px.Format(font_color="#9C0006", background_color="#FFC7CE")


def rules(reopen, path):
    """{sqref: [rule, ...]} as read by openpyxl."""
    sh = reopen(path)["Sheet1"]
    return {str(cf.sqref): list(cf.rules) for cf in sh.conditional_formatting}


def save(out, *pairs):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_column("A1", [1, 5, 10, 15, 20])
    for rng, cf in pairs:
        ws.add_conditional_format(rng, cf)
    path = out()
    wb.save(path)
    return path


def test_cell_rules(out, reopen):
    path = save(
        out,
        ("A1:A5", px.ConditionalFormatCell(rule=CellRule.greater_than(5), format=px.Format(bold=True))),
        ("B1:B5", px.ConditionalFormatCell(rule=CellRule.between(1, 10), format=RED)),
        ("C1:C5", px.ConditionalFormatCell(rule=CellRule.equal_to("abc"))),
        ("D1:D5", px.ConditionalFormatCell(rule=CellRule.less_than("=$Z$1"))),
        ("E1:E5", px.ConditionalFormatCell(rule=CellRule.not_between(dt.date(2024, 1, 1), px.Formula("=$Z$2")))),
        ("F1:F5", px.ConditionalFormatCell(rule=CellRule.not_equal_to(True))),
        ("G1:G5", px.ConditionalFormatCell(rule=CellRule.greater_than_or_equal_to(0.5))),
        ("H1:H5", px.ConditionalFormatCell(rule=CellRule.less_than_or_equal_to(dt.time(12, 0)))),
    )
    got = rules(reopen, path)
    (a,) = got["A1:A5"]
    assert (a.type, a.operator, a.formula) == ("cellIs", "greaterThan", ["5"])
    assert a.dxf.font.b
    (b,) = got["B1:B5"]
    assert (b.operator, b.formula) == ("between", ["1", "10"])
    assert b.dxf.fill.bgColor.rgb.endswith("FFC7CE")
    assert got["C1:C5"][0].formula == ['"abc"']
    assert got["D1:D5"][0].formula == ["$Z$1"]
    assert (got["E1:E5"][0].operator, got["E1:E5"][0].formula) == ("notBetween", ["45292", "$Z$2"])
    assert (got["F1:F5"][0].operator, got["F1:F5"][0].formula) == ("notEqual", ["TRUE"])
    assert got["G1:G5"][0].operator == "greaterThanOrEqual"
    assert (got["H1:H5"][0].operator, got["H1:H5"][0].formula) == ("lessThanOrEqual", ["0.5"])


def test_text_date_average_top(out, reopen):
    TR = px.ConditionalFormatTextRule
    path = save(
        out,
        ("A1:A5", px.ConditionalFormatText(rule=TR.contains("foo"), format=RED)),
        ("B1:B5", px.ConditionalFormatText(rule=TR.does_not_contain("foo"))),
        ("C1:C5", px.ConditionalFormatText(rule=TR.begins_with("ab"))),
        ("D1:D5", px.ConditionalFormatText(rule=TR.ends_with("yz"))),
        ("E1:E5", px.ConditionalFormatDate(rule=px.ConditionalFormatDateRule.Last7Days)),
        ("F1:F5", px.ConditionalFormatDate()),
        ("G1:G5", px.ConditionalFormatAverage(rule=px.ConditionalFormatAverageRule.BelowAverage)),
        ("H1:H5", px.ConditionalFormatAverage(rule=px.ConditionalFormatAverageRule.TwoStandardDeviationsAbove)),
        ("I1:I5", px.ConditionalFormatTop()),
        ("J1:J5", px.ConditionalFormatTop(rule=px.ConditionalFormatTopRule.bottom_percent(25))),
    )
    got = {k: v[0] for k, v in rules(reopen, path).items()}
    assert (got["A1:A5"].type, got["A1:A5"].operator, got["A1:A5"].text) == ("containsText", "containsText", "foo")
    assert got["B1:B5"].type == "notContainsText"
    assert (got["C1:C5"].type, got["C1:C5"].text) == ("beginsWith", "ab")
    assert (got["D1:D5"].type, got["D1:D5"].text) == ("endsWith", "yz")
    assert (got["E1:E5"].type, got["E1:E5"].timePeriod) == ("timePeriod", "last7Days")
    assert got["F1:F5"].timePeriod == "yesterday"
    assert (got["G1:G5"].type, got["G1:G5"].aboveAverage) == ("aboveAverage", False)
    assert got["H1:H5"].stdDev == 2
    assert (got["I1:I5"].type, got["I1:I5"].rank, got["I1:I5"].bottom) == ("top10", 10, None)
    assert (got["J1:J5"].rank, got["J1:J5"].percent, got["J1:J5"].bottom) == (25, True, True)


def test_duplicate_blank_error_formula(out, reopen):
    path = save(
        out,
        ("A1:A5", px.ConditionalFormatDuplicate(format=RED)),
        ("B1:B5", px.ConditionalFormatDuplicate(invert=True)),
        ("C1:C5", px.ConditionalFormatBlank()),
        ("D1:D5", px.ConditionalFormatBlank(invert=True)),
        ("E1:E5", px.ConditionalFormatError()),
        ("F1:F5", px.ConditionalFormatError(invert=True)),
        ("G1:G5", px.ConditionalFormatFormula(rule="=ISODD(G1)", format=RED)),
        ("H1:H5", px.ConditionalFormatFormula(rule=px.Formula("=$A1>5"))),
    )
    got = {k: v[0] for k, v in rules(reopen, path).items()}
    assert got["A1:A5"].type == "duplicateValues"
    assert got["B1:B5"].type == "uniqueValues"
    assert got["C1:C5"].type == "containsBlanks"
    assert got["D1:D5"].type == "notContainsBlanks"
    assert got["E1:E5"].type == "containsErrors"
    assert got["F1:F5"].type == "notContainsErrors"
    assert (got["G1:G5"].type, got["G1:G5"].formula) == ("expression", ["ISODD(G1)"])
    assert got["H1:H5"].formula == ["$A1>5"]


def test_multi_range_and_stop_if_true(out, reopen):
    cf = px.ConditionalFormatCell(
        rule=CellRule.greater_than(1), format=RED, multi_range="A1:A5 $C$1:$C$5", stop_if_true=True
    )
    path = save(out, ("A1:A5", cf))
    got = rules(reopen, path)
    (rule,) = got["A1:A5 C1:C5"]
    assert rule.stopIfTrue


def test_color_scales(out, reopen):
    path = save(
        out,
        ("A1:A5", px.ConditionalFormat2ColorScale()),
        (
            "B1:B5",
            px.ConditionalFormat2ColorScale(min_color="#FF0000", max_color=px.Color.Green)
            .set_minimum(T.Number, 3)
            .set_maximum(T.Percentile, 90),
        ),
        ("C1:C5", px.ConditionalFormat3ColorScale()),
        (
            "D1:D5",
            px.ConditionalFormat3ColorScale(
                min_type=T.Percent,
                min_value=10,
                mid_type=T.Number,
                mid_value=7.5,
                max_type=T.Formula,
                max_value="=$Z$1",
                min_color=(0, 0, 255),
                mid_color="white",
                max_color=0xFF0000,
            ),
        ),
    )
    got = {k: v[0] for k, v in rules(reopen, path).items()}
    a = got["A1:A5"].colorScale
    assert [c.type for c in a.cfvo] == ["min", "max"]
    assert [c.rgb for c in a.color] == ["FFFFEF9C", "FF63BE7B"]
    b = got["B1:B5"].colorScale
    assert [(c.type, c.val) for c in b.cfvo] == [("num", 3), ("percentile", 90)]
    assert [c.rgb for c in b.color] == ["FFFF0000", "FF008000"]
    c = got["C1:C5"].colorScale
    assert [(v.type, v.val) for v in c.cfvo] == [("min", 0), ("percentile", 50), ("max", 0)]
    d = got["D1:D5"].colorScale
    assert [(v.type, v.val) for v in d.cfvo] == [("percent", 10), ("num", 7.5), ("formula", "$Z$1")]
    assert [v.rgb for v in d.color] == ["FF0000FF", "FFFFFFFF", "FFFF0000"]


def test_color_scale_properties():
    cf = px.ConditionalFormat3ColorScale()
    assert (cf.min_type, cf.mid_type, cf.max_type) == (T.Lowest, T.Percentile, T.Highest)
    assert cf.min_value is None and cf.mid_value == 50 and cf.max_value is None
    assert cf.min_color is None
    cf.min_color = "#123456"
    assert cf.min_color == px.Color("#123456")
    cf.set_midpoint(T.Number, dt.date(2024, 1, 1))
    assert cf.mid_value == dt.date(2024, 1, 1)
    assert cf.copy() == cf
    assert cf.set(stop_if_true=True) is cf and cf.stop_if_true


def test_data_bars(out, reopen, part):
    full = px.ConditionalFormatDataBar(
        min_type=T.Number,
        min_value=1,
        max_type=T.Percent,
        max_value=90,
        fill_color="#FF0000",
        negative_fill_color="#0000FF",
        solid_fill=True,
        direction=px.ConditionalFormatDataBarDirection.RightToLeft,
        bar_only=True,
        axis_position=px.ConditionalFormatDataBarAxisPosition.Midpoint,
        axis_color="#123456",
        border_color="#00FF00",
        negative_border_color="#FFFF00",
    )
    path = save(
        out,
        ("A1:A5", px.ConditionalFormatDataBar()),
        ("B1:B5", full),
        (
            "C1:C5",
            px.ConditionalFormatDataBar(border_off=True, axis_position=px.ConditionalFormatDataBarAxisPosition.None_),
        ),
        ("D1:D5", px.ConditionalFormatDataBar(classic_style=True)),
    )
    got = {k: v[0] for k, v in rules(reopen, path).items()}
    assert got["A1:A5"].type == "dataBar"
    b = got["B1:B5"].dataBar
    assert [(c.type, c.val) for c in b.cfvo] == [("num", 1), ("percent", 90)]
    assert b.color.rgb == "FFFF0000" and b.showValue is False
    assert [c.type for c in got["D1:D5"].dataBar.cfvo] == ["min", "max"]
    xml = part(path, "xl/worksheets/sheet1.xml")
    ext = xml[xml.index("<x14:conditionalFormattings") :]
    assert 'gradient="0"' in ext and 'direction="rightToLeft"' in ext and 'axisPosition="middle"' in ext
    assert '<x14:negativeFillColor rgb="FF0000FF"/>' in ext
    assert '<x14:borderColor rgb="FF00FF00"/>' in ext
    assert '<x14:negativeBorderColor rgb="FFFFFF00"/>' in ext
    assert '<x14:axisColor rgb="FF123456"/>' in ext
    b_rule = ext.split("<xm:sqref>B1:B5")[0].rsplit("<x14:cfRule", 1)[1]
    assert 'border="1"' in b_rule
    c_rule = ext.split("<xm:sqref>C1:C5")[0].rsplit("<x14:cfRule", 1)[1]
    assert 'axisPosition="none"' in c_rule and 'border="1"' not in c_rule
    assert xml.count("<x14:cfRule") == 3  # the classic bar has no extension


def test_icon_sets(out, reopen, part):
    icon_type = px.ConditionalFormatIconType
    icons = [
        px.ConditionalFormatCustomIcon(),
        px.ConditionalFormatCustomIcon(rule_type=T.Percentile, value=50),
        px.ConditionalFormatCustomIcon(rule_type=T.Number, value=90, greater_than=True),
    ]
    path = save(
        out,
        ("A1:A5", px.ConditionalFormatIconSet()),
        ("B1:B5", px.ConditionalFormatIconSet(icon_type=icon_type.FourArrows, reverse_icons=True, icons_only=True)),
        ("C1:C5", px.ConditionalFormatIconSet(icon_type=icon_type.ThreeFlags, icons=icons)),
        ("D1:D5", px.ConditionalFormatIconSet(icon_type=icon_type.FiveBoxes)),
        (
            "E1:E5",
            px.ConditionalFormatIconSet(
                icon_type=icon_type.ThreeArrows,
                icons=[
                    px.ConditionalFormatCustomIcon(),
                    px.ConditionalFormatCustomIcon(value=40, icon_type=icon_type.ThreeStars, icon_index=2),
                    px.ConditionalFormatCustomIcon(value=80, no_icon=True),
                ],
            ),
        ),
    )
    got = {k: v[0] for k, v in rules(reopen, path).items()}
    a = got["A1:A5"].iconSet
    assert a.iconSet is None  # 3TrafficLights1 is the default and is not written
    assert [(c.type, c.val) for c in a.cfvo] == [("percent", 0), ("percent", 33), ("percent", 67)]
    b = got["B1:B5"].iconSet
    assert b.iconSet == "4Arrows" and b.reverse and b.showValue is False and len(b.cfvo) == 4
    c = got["C1:C5"].iconSet
    assert c.iconSet == "3Flags"
    assert [(v.type, v.val, v.gte) for v in c.cfvo] == [
        ("percent", 0, None),
        ("percentile", 50, None),
        ("num", 90, False),
    ]
    xml = part(path, "xl/worksheets/sheet1.xml")
    assert 'iconSet="5Boxes"' in xml and 'custom="1"' in xml
    assert '<x14:cfIcon iconSet="3Stars" iconId="2"/>' in xml
    assert '<x14:cfIcon iconSet="NoIcons" iconId="0"/>' in xml


def test_custom_icon_properties():
    icon = px.ConditionalFormatCustomIcon()
    assert icon.rule_type == T.Percent and icon.value == 0 and not icon.greater_than
    assert icon.icon_type is None and icon.icon_index == 0 and not icon.no_icon
    icon.value = px.Formula("=$A$1")
    icon.rule_type = T.Formula
    assert isinstance(icon.value, px.Formula) and icon.value.formula == "=$A$1"
    s = px.ConditionalFormatIconSet(icons=[icon, icon, icon])
    assert len(s.icons) == 3 and s.icons[0] == icon
    with pytest.raises(px.ParameterError, match="icon_index"):
        px.ConditionalFormatCustomIcon(icon_type=px.ConditionalFormatIconType.ThreeFlags, icon_index=3)
    with pytest.raises(px.ParameterError, match="rule_type"):
        px.ConditionalFormatCustomIcon(rule_type=T.Highest)
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.ParameterError, match="custom icons"):
        ws.add_conditional_format("A1:A5", px.ConditionalFormatIconSet(icons=[icon, icon]))


def test_rule_objects():
    rule = CellRule.between(1, "=B1")
    assert rule.criteria == "between" and rule.values == (1, "=B1")
    assert repr(rule) == "ConditionalFormatCellRule.between(1, '=B1')"
    assert CellRule.equal_to(1) == CellRule.equal_to(1)
    text = px.ConditionalFormatTextRule.begins_with("a")
    assert (text.criteria, text.value) == ("begins_with", "a")
    assert repr(text) == 'ConditionalFormatTextRule.begins_with("a")'
    top = px.ConditionalFormatTopRule.top_percent(5)
    assert (top.criteria, top.value) == ("top_percent", 5)
    assert px.ConditionalFormatTop().rule == px.ConditionalFormatTopRule.top(10)
    with pytest.raises(px.ParameterError, match="1..=1000"):
        px.ConditionalFormatTopRule.top(0)
    with pytest.raises(px.ParameterError, match="1..=100"):
        px.ConditionalFormatTopRule.bottom_percent(101)


def test_common_properties():
    cf = px.ConditionalFormatCell()
    assert cf.rule is None and cf.format is None and cf.multi_range is None and not cf.stop_if_true
    cf.format = px.Format(bold=True)
    assert cf.format == px.Format(bold=True)
    cf.rule = CellRule.less_than(3)
    clone = cf.copy()
    assert clone == cf
    clone.stop_if_true = True
    assert clone != cf
    assert px.ConditionalFormatDuplicate(invert=True).invert
    assert px.ConditionalFormatFormula(rule=px.Formula("=A1")).rule == "=A1"
    assert px.ConditionalFormatDate().rule == px.ConditionalFormatDateRule.Yesterday
    assert px.ConditionalFormatAverage().rule == px.ConditionalFormatAverageRule.AboveAverage
    bar = px.ConditionalFormatDataBar()
    assert bar.min_type == T.Automatic and bar.direction == px.ConditionalFormatDataBarDirection.Context
    assert bar.axis_position == px.ConditionalFormatDataBarAxisPosition.Automatic
    assert not (bar.solid_fill or bar.border_off or bar.bar_only or bar.classic_style)


def test_legacy_methods_still_work(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    two = px.ConditionalFormat2ColorScale()
    two.min_color = "red"
    two.max_color = px.Color.Blue
    three = px.ConditionalFormat3ColorScale()
    three.mid_color = "#FFFF00"
    bar = px.ConditionalFormatDataBar()
    bar.fill_color = "#00FF00"
    ws.add_conditional_format_2color_scale(0, 0, 4, 0, two)
    ws.add_conditional_format_3color_scale(0, 1, 4, 1, three)
    ws.add_conditional_format_data_bar("C1:C5", cond_fmt=bar)
    path = out()
    wb.save(path)
    got = rules(reopen, path)
    assert [c.rgb for c in got["A1:A5"][0].colorScale.color] == ["FFFF0000", "FF0000FF"]
    assert got["B1:B5"][0].colorScale.color[1].rgb == "FFFFFF00"
    assert got["C1:C5"][0].dataBar.color.rgb == "FF00FF00"
    with pytest.raises(TypeError, match="expects a ConditionalFormatDataBar"):
        ws.add_conditional_format_data_bar("D1:D5", two)  # pyrefly: ignore[bad-argument-type]


def test_errors():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.ParameterError, match="needs a rule"):
        ws.add_conditional_format("A1:A5", px.ConditionalFormatCell())
    with pytest.raises(px.ParameterError, match="needs a rule"):
        ws.add_conditional_format("A1:A5", px.ConditionalFormatText())
    with pytest.raises(px.ParameterError, match="formula"):
        ws.add_conditional_format("A1:A5", px.ConditionalFormatFormula())
    with pytest.raises(TypeError, match="ConditionalFormat"):
        ws.add_conditional_format("A1:A5", px.Format())  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.RangeError):
        ws.add_conditional_format("B5:A1", px.ConditionalFormatBlank())
    with pytest.raises(px.ParameterError, match="min_type"):
        px.ConditionalFormat2ColorScale(min_type=T.Highest)
    with pytest.raises(px.ParameterError, match="mid_type"):
        px.ConditionalFormat3ColorScale(mid_type=T.Lowest)
    with pytest.raises(px.ParameterError, match="max_type"):
        px.ConditionalFormatDataBar(max_type=T.Lowest)
    with pytest.raises(px.ParameterError, match="not str"):
        px.ConditionalFormat2ColorScale(min_type=T.Number, min_value="abc")
    with pytest.raises(px.ParameterError, match="0..=100"):
        px.ConditionalFormatDataBar().set_minimum(T.Percent, 150)
    with pytest.raises(px.ParameterError, match="must be set"):
        ws.add_conditional_format("A1:A5", px.ConditionalFormat2ColorScale(min_type=T.Number))
    with pytest.raises(px.ParameterError, match="multi_range"):
        px.ConditionalFormatBlank(multi_range=" ")
    with pytest.raises(TypeError):
        px.ConditionalFormatCell(rule=5)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.ConditionalFormatDataBar(colour="red")  # pyrefly: ignore[unexpected-keyword]
    with pytest.raises(px.ParameterError):
        px.ConditionalFormatDataBar(fill_color="#GG0000")
