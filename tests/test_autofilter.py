"""Autofilters: Worksheet.autofilter/filter_column/filter_automatic_off, FilterCondition."""

import pytest

import pyrust_xlsxwriter as px

ROWS = [
    ["Region", "Item", "Volume"],
    ["East", "Apple", 9000],
    ["West", "Pear", 5000],
    ["", "Grape", 3000],
    ["East", "Pear", 7000],
    ["South", "Apple", 2000],
]


def sheet():
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_rows("A1", ROWS)
    return wb, ws


def test_autofilter_range(out, reopen):
    wb, ws = sheet()
    ws.autofilter("A1:C6")
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    assert sh.auto_filter.ref == "A1:C6"
    assert sh.auto_filter.filterColumn == []


def test_autofilter_ints(out, reopen):
    wb, ws = sheet()
    ws.autofilter(0, 0, 5, 2)
    path = out()
    wb.save(path)
    assert reopen(path)["Sheet1"].auto_filter.ref == "A1:C6"


def test_list_filter_hides_rows(out, reopen):
    wb, ws = sheet()
    ws.autofilter("A1:C6")
    ws.filter_column("A", px.FilterCondition().add_list_filter("East").add_list_filter("West"))
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    (col,) = sh.auto_filter.filterColumn
    assert col.colId == 0
    assert sorted(col.filters.filter) == ["East", "West"]
    assert not col.filters.blank
    hidden = {r for r in range(1, 7) if sh.row_dimensions[r].hidden}
    assert hidden == {4, 6}


def test_list_filter_blanks_and_numbers(out, reopen):
    wb, ws = sheet()
    ws.autofilter("A1:C6")
    ws.filter_column(0, px.FilterCondition(list_filters=["East"], blanks=True))
    ws.filter_column(2, px.FilterCondition(list_filters=[9000, 2000.0]))
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    cols = {c.colId: c for c in sh.auto_filter.filterColumn}
    assert cols[0].filters.blank
    assert cols[0].filters.filter == ["East"]
    assert sorted(cols[2].filters.filter) == ["2000", "9000"]


def test_custom_filters(out, reopen):
    wb, ws = sheet()
    ws.autofilter("A1:C6")
    between = (
        px.FilterCondition()
        .add_custom_filter(px.FilterCriteria.GreaterThanOrEqualTo, 3000)
        .add_custom_filter(px.FilterCriteria.LessThanOrEqualTo, 8000)
    )
    ws.filter_column("C", between)
    either = px.FilterCondition(
        custom_filters=[(px.FilterCriteria.BeginsWith, "Ap"), (px.FilterCriteria.EndsWith, "pe")],
        boolean_or=True,
    )
    ws.filter_column("B", either)
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    cols = {c.colId: c for c in sh.auto_filter.filterColumn}
    vol = cols[2].customFilters
    assert vol._and
    assert [(f.operator, f.val) for f in vol.customFilter] == [
        ("greaterThanOrEqual", "3000"),
        ("lessThanOrEqual", "8000"),
    ]
    item = cols[1].customFilters
    assert not item._and
    assert [(f.operator, f.val) for f in item.customFilter] == [("equal", "Ap*"), ("equal", "*pe")]
    hidden = {r for r in range(1, 7) if sh.row_dimensions[r].hidden}
    # 9000 and 2000 are outside 3000..8000; "Pear" neither begins with "Ap" nor ends with "pe".
    assert hidden == {2, 3, 5, 6}


def test_filter_automatic_off(out, reopen):
    wb, ws = sheet()
    ws.autofilter("A1:C6")
    ws.filter_automatic_off()
    ws.filter_column("A", px.FilterCondition(list_filters=["East"]))
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    assert not any(sh.row_dimensions[r].hidden for r in range(1, 7))
    assert sh.auto_filter.filterColumn[0].filters.filter == ["East"]


def test_filter_condition_properties():
    fc = px.FilterCondition()
    assert fc.list_filters == [] and not fc.blanks and fc.custom_filters == [] and not fc.boolean_or
    fc.add_list_filter("a").add_list_filter(2).add_list_blanks_filter()
    assert fc.list_filters == ["a", 2] and fc.blanks
    fc2 = px.FilterCondition(custom_filters=[(px.FilterCriteria.Contains, "x")]).add_custom_boolean_or()
    assert fc2.custom_filters == [(px.FilterCriteria.Contains, "x")] and fc2.boolean_or
    assert fc.copy() == fc and fc != fc2


def test_filter_errors():
    _wb, ws = sheet()
    with pytest.raises(px.ParameterError):  # no autofilter range yet
        ws.filter_column(0, px.FilterCondition(list_filters=["East"]))
    ws.autofilter("A1:B6")
    with pytest.raises(px.ParameterError):  # outside the autofilter columns
        ws.filter_column("D", px.FilterCondition(list_filters=["East"]))
    with pytest.raises(px.ParameterError, match="no conditions"):
        ws.filter_column(0, px.FilterCondition())
    with pytest.raises(px.ParameterError, match="single column"):
        ws.filter_column("A:B", px.FilterCondition(list_filters=["x"]))
    fc = (
        px.FilterCondition()
        .add_custom_filter(px.FilterCriteria.EqualTo, 1)
        .add_custom_filter(px.FilterCriteria.EqualTo, 2)
    )
    with pytest.raises(px.ParameterError, match="at most 2"):
        fc.add_custom_filter(px.FilterCriteria.EqualTo, 3)
    assert len(fc.custom_filters) == 2  # unchanged after the error
    with pytest.raises(px.ParameterError, match="not both"):
        px.FilterCondition(list_filters=["a"]).add_custom_filter(px.FilterCriteria.EqualTo, 1)
    with pytest.raises(TypeError):
        px.FilterCondition().add_list_filter(True)
    with pytest.raises(TypeError):
        ws.filter_column(0, "East")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.RangeError):
        ws.autofilter("C1:A6")
