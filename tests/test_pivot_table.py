"""Pivot tables: PivotTable, PivotTableValueField, PivotTableFunction/Layout/Style."""

import re
import zipfile

import pytest

import pyrust_xlsxwriter as px

ROWS = [
    ["Region", "Item", "Volume", "Month"],
    ["East", "Apple", 9000, "July"],
    ["West", "Pear", 5000, "July"],
    ["North", "Apple", 3000, "August"],
    ["East", "Grape", 8000, "August"],
]


def book():
    wb = px.Workbook()
    ws = wb.add_worksheet("Data")
    ws.write_rows("A1", ROWS)
    return wb, ws


def pivot_xml(part, path, n=1):
    return part(path, f"xl/pivotTables/pivotTable{n}.xml")


def test_basic_pivot_table(out, reopen, part):
    wb, _ = book()
    pivot = px.PivotTable(data_source="Data!A1:D5", row_fields=["Region"], value_fields=["Volume"])
    wb.add_worksheet("Pivot").add_pivot_table("A3", pivot)
    path = out()
    wb.save(path)
    reopen(path)
    xml = pivot_xml(part, path)
    assert '<location ref="A3"' in xml
    assert '<rowFields count="1"><field x="0"/></rowFields>' in xml
    assert '<dataField name="Sum of Volume" fld="2"' in xml
    assert 'name="PivotStyleLight16"' in xml
    assert 'rowGrandTotals="0"' not in xml and 'compact="0"' not in xml
    cache = part(path, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert '<worksheetSource ref="A1:D5" sheet="Data"/>' in cache
    assert re.findall(r'cacheField name="(\w+)"', cache) == ["Region", "Item", "Volume", "Month"]
    with zipfile.ZipFile(path) as zf:
        assert "xl/pivotTables/pivotTable1.xml" in zf.namelist()


def test_all_options(out, reopen, part):
    wb, _ = book()
    pivot = (
        px.PivotTable(
            data_source=("Data", 0, 0, 4, 3),
            style=px.PivotTableStyle.Medium4,
            layout=px.PivotTableLayout.Tabular,
            row_grand_totals=False,
            column_grand_totals=False,
        )
        .add_row_field("Region")
        .add_column_field("Item")
        .add_filter_field("Month")
        .add_value_field(
            px.PivotTableValueField(
                "Volume", function=px.PivotTableFunction.Average, custom_name="Avg", num_format="#,##0"
            )
        )
        .add_value_field("Volume")
    )
    wb.add_worksheet("Pivot").add_pivot_table(2, 0, pivot)
    path = out()
    wb.save(path)
    reopen(path)
    xml = pivot_xml(part, path)
    assert 'rowGrandTotals="0"' in xml and 'colGrandTotals="0"' in xml
    assert 'compact="0"' in xml and 'outline="0"' in xml  # tabular layout
    assert '<location ref="A5"' in xml  # moved down below the filter field
    assert '<pageFields count="1"><pageField fld="3"' in xml
    assert '<colFields count="2"><field x="1"/><field x="-2"/></colFields>' in xml
    assert '<dataField name="Avg" fld="2" subtotal="average"' in xml and 'numFmtId="164"' in xml
    assert 'name="PivotStyleMedium4"' in xml
    styles = part(path, "xl/styles.xml")
    assert 'formatCode="#,##0"' in styles


def test_outline_layout_style_none_and_functions(out, part):
    wb, ws = book()
    funcs = [
        (px.PivotTableFunction.Count, "count"),
        (px.PivotTableFunction.Max, "max"),
        (px.PivotTableFunction.Min, "min"),
        (px.PivotTableFunction.Product, "product"),
        (px.PivotTableFunction.CountNumbers, "countNums"),
        (px.PivotTableFunction.StdDev, "stdDev"),
        (px.PivotTableFunction.StdDevP, "stdDevp"),
        (px.PivotTableFunction.Var, "var"),
        (px.PivotTableFunction.VarP, "varp"),
    ]
    pivot = px.PivotTable(
        data_source="A1:D5",  # same sheet: the sheet name is filled in
        layout=px.PivotTableLayout.Outline,
        style=px.PivotTableStyle.None_,
        row_fields=["Region"],
        value_fields=[
            px.PivotTableValueField("Volume", function=f, custom_name=f"V{i}") for i, (f, _) in enumerate(funcs)
        ],
    )
    ws.add_pivot_table("H2", pivot)
    path = out()
    wb.save(path)
    xml = pivot_xml(part, path)
    for _, name in funcs:
        assert f'subtotal="{name}"' in xml
    assert 'compact="0" outline="1"' in xml  # outline layout
    (style_info,) = re.findall(r"<pivotTableStyleInfo[^>]*>", xml)
    assert "name=" not in style_info  # PivotTableStyle.None_
    cache = part(path, "xl/pivotCache/pivotCacheDefinition1.xml")
    assert 'sheet="Data"' in cache


def test_field_names(out, part):
    wb, _ = book()
    pivot = px.PivotTable(data_source="Data!A1:D5", field_names=["Region", "Item", "Volume", "Month"])
    pivot.add_row_field("Item").add_value_field("Volume")
    wb.add_worksheet().add_pivot_table("A1", pivot)
    path = out()
    wb.save(path)
    assert '<rowFields count="1"><field x="1"/></rowFields>' in pivot_xml(part, path)


def test_properties_round_trip():
    pivot = px.PivotTable()
    assert pivot.data_source is None and pivot.field_names is None
    assert pivot.row_fields == [] and pivot.column_fields == [] and pivot.filter_fields == []
    assert pivot.value_fields == []
    assert pivot.style == px.PivotTableStyle.Light16 and pivot.layout == px.PivotTableLayout.Compact
    assert pivot.row_grand_totals and pivot.column_grand_totals
    pivot.add_row_field("A").add_column_field("B").add_filter_field("C").add_value_field("D")
    assert (pivot.row_fields, pivot.column_fields, pivot.filter_fields) == (["A"], ["B"], ["C"])
    (vf,) = pivot.value_fields
    assert vf.field == "D" and vf.function == px.PivotTableFunction.Sum
    assert vf.custom_name is None and vf.num_format is None
    pivot.data_source = ("My Data", 0, 0, 9, 3)
    assert pivot.data_source == "'My Data'!$A$1:$D$10"
    clone = pivot.copy()
    assert clone == pivot
    clone.layout = px.PivotTableLayout.Outline
    assert clone != pivot
    vf2 = px.PivotTableValueField("X", custom_name="Total X")
    vf2.function = px.PivotTableFunction.Max
    vf2.num_format = "0.0"
    assert (vf2.field, vf2.function, vf2.custom_name, vf2.num_format) == (
        "X",
        px.PivotTableFunction.Max,
        "Total X",
        "0.0",
    )
    assert vf2.copy() == vf2
    assert len([n for n in dir(px.PivotTableStyle) if n[0].isupper()]) == 85


def test_errors(out):
    wb, ws = book()
    other = wb.add_worksheet("P")
    with pytest.raises(px.ParameterError, match="no data source"):
        other.add_pivot_table("A1", px.PivotTable(row_fields=["Region"]))
    with pytest.raises(px.ParameterError, match="valid range"):
        other.add_pivot_table("A1", px.PivotTable(data_source="Data!nonsense"))
    with pytest.raises(px.ParameterError, match="header row"):
        other.add_pivot_table("A1", px.PivotTable(data_source="Data!A1:D1"))
    with pytest.raises(px.ParameterError, match="more than once"):
        other.add_pivot_table(
            "A1", px.PivotTable(data_source="Data!A1:D5", row_fields=["Region"], column_fields=["Region"])
        )
    with pytest.raises(px.ParameterError, match="field names"):
        other.add_pivot_table("A1", px.PivotTable(data_source="Data!A1:D5", field_names=["a"]))
    with pytest.raises(px.ParameterError, match="overlaps its source"):
        ws.add_pivot_table("B2", px.PivotTable(data_source="A1:D5", value_fields=["Volume"]))
    with pytest.raises(px.ParameterError):
        px.PivotTable().add_row_field("")
    with pytest.raises(px.ParameterError):
        px.PivotTableValueField("")
    with pytest.raises(TypeError):
        px.PivotTable(value_fields=[1])  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        px.PivotTable(value_fields="Volume")
    with pytest.raises(TypeError):
        other.add_pivot_table("A1", "pivot")  # pyrefly: ignore[bad-argument-type]
    # A field that isn't in the header row is reported when the file is saved.
    other.add_pivot_table("A1", px.PivotTable(data_source="Data!A1:D5", row_fields=["Nope"]))
    with pytest.raises(px.ParameterError, match="Nope"):
        wb.save(out())
