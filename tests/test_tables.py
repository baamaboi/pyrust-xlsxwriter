"""Worksheet tables: Table, TableColumn, TableStyle, TableFunction."""

import pytest

import pyrust_xlsxwriter as px

DATA = [
    ["Apples", 10000, 5000, 8000],
    ["Pears", 2000, 3000, 4000],
    ["Bananas", 6000, 6000, 6500],
]


def write_data(ws, row=1, col=1):
    for i, values in enumerate(DATA):
        ws.write_row(row + i, col, values)


def test_default_table(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    write_data(ws)
    ws.add_table("B1:E4")
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    ((name, ref),) = sh.tables.items()
    assert ref == "B1:E4"
    table = sh.tables[name]
    assert table.tableStyleInfo.name == "TableStyleMedium9"
    assert table.tableStyleInfo.showRowStripes
    assert not table.tableStyleInfo.showColumnStripes
    assert [c.name for c in table.tableColumns] == ["Column1", "Column2", "Column3", "Column4"]
    assert sh["B1"].value == "Column1"
    assert table.autoFilter is not None


def test_table_properties_round_trip():
    t = px.Table()
    assert t.header_row and t.banded_rows and t.autofilter
    assert not (t.total_row or t.banded_columns or t.first_column or t.last_column)
    assert t.style == px.TableStyle.Medium9
    assert t.name is None and t.alt_text is None and t.columns == []
    t = px.Table(
        header_row=True,
        total_row=True,
        banded_rows=False,
        banded_columns=True,
        first_column=True,
        last_column=True,
        autofilter=False,
        name="Sales",
        style=px.TableStyle.Light11,
        alt_text="Sales table",
        alt_text_title="Sales",
        columns=["A", px.TableColumn("B")],
    )
    assert t.total_row and t.banded_columns and t.first_column and t.last_column
    assert not t.banded_rows and not t.autofilter
    assert t.name == "Sales" and t.style == px.TableStyle.Light11
    assert t.alt_text == "Sales table" and t.alt_text_title == "Sales"
    assert [c.header for c in t.columns] == ["A", "B"]
    t2 = t.copy()
    assert t2 == t
    t2.name = "Other"
    assert t2 != t and t.name == "Sales"
    assert t.set(banded_rows=True) is t and t.banded_rows


def test_table_options_written(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    write_data(ws, row=1)
    money = px.Format(num_format="#,##0")
    table = px.Table(
        name="Fruit",
        style=px.TableStyle.Dark3,
        total_row=True,
        banded_rows=False,
        banded_columns=True,
        first_column=True,
        last_column=True,
        autofilter=False,
        columns=[
            px.TableColumn("Product", total_label="Totals", header_format=px.Format(bold=True)),
            px.TableColumn("Q1", total_function=px.TableFunction.Sum, format=money),
            px.TableColumn("Q2", total_function=px.TableFunction.Average),
            px.TableColumn("Q3", total_function=px.TableFunction.custom("=SUM([Q3])/2")),
            px.TableColumn("Total", formula="=SUM(Fruit[@[Q1]:[Q3]])"),
        ],
    )
    ws.add_table(0, 0, 4, 4, table)
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    t = sh.tables["Fruit"]
    assert t.ref == "A1:E5"
    assert t.totalsRowCount == 1
    assert t.autoFilter is None
    info = t.tableStyleInfo
    assert info.name == "TableStyleDark3"
    assert not info.showRowStripes and info.showColumnStripes
    assert info.showFirstColumn and info.showLastColumn
    cols = t.tableColumns
    assert [c.name for c in cols] == ["Product", "Q1", "Q2", "Q3", "Total"]
    assert cols[0].totalsRowLabel == "Totals"
    assert cols[1].totalsRowFunction == "sum"
    assert cols[2].totalsRowFunction == "average"
    assert cols[3].totalsRowFunction == "custom"
    assert cols[4].calculatedColumnFormula is not None
    assert sh["A1"].value == "Product" and sh["A1"].font.b
    assert sh["A5"].value == "Totals"
    assert sh["B5"].value == "=SUBTOTAL(109,[Q1])"
    assert sh["B2"].number_format == "#,##0"
    assert sh["E2"].value.startswith("=SUM(")


@pytest.mark.parametrize(
    ("function", "subtotal"),
    [
        (px.TableFunction.Average, 101),
        (px.TableFunction.CountNumbers, 102),
        (px.TableFunction.Count, 103),
        (px.TableFunction.Max, 104),
        (px.TableFunction.Min, 105),
        (px.TableFunction.StdDev, 107),
        (px.TableFunction.Sum, 109),
        (px.TableFunction.Var, 110),
    ],
)
def test_total_functions(out, reopen, function, subtotal):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_column("A2", [1, 2, 3])
    ws.add_table("A1:A5", px.Table(total_row=True, columns=[px.TableColumn("N", total_function=function)]))
    path = out()
    wb.save(path)
    assert reopen(path)["Sheet1"]["A5"].value == f"=SUBTOTAL({subtotal},[N])"


def test_table_function_values():
    assert px.TableFunction.Sum == px.TableFunction.Sum
    assert px.TableFunction.Sum != px.TableFunction.Max
    assert px.TableFunction.Sum.name == "Sum" and px.TableFunction.Sum.formula is None
    assert px.TableFunction.None_.name == "None_"
    custom = px.TableFunction.custom(px.Formula("=MAX([N])"))
    assert custom.name == "Custom" and custom.formula == "=MAX([N])"
    assert custom == px.TableFunction.custom("=MAX([N])")
    assert repr(custom) == 'TableFunction.custom("=MAX([N])")'
    assert len({px.TableFunction.Sum, px.TableFunction.Sum, custom}) == 2
    with pytest.raises(px.ParameterError):
        px.TableFunction.custom("=")


def test_table_column_properties():
    col = px.TableColumn()
    assert col.header is None and col.total_function is None and col.format is None
    col = px.TableColumn("H", total_label="T", formula=px.Formula("=1"), format=px.Format(italic=True))
    assert col.header == "H" and col.total_label == "T" and col.formula == "=1"
    assert col.format == px.Format(italic=True) and col.header_format is None
    col.total_function = px.TableFunction.Max
    assert col.total_function == px.TableFunction.Max
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.TableColumn(bogus=1)  # pyrefly: ignore[unexpected-keyword]


def test_all_table_styles_exist(out, reopen):
    names = [n for n in dir(px.TableStyle) if n[0].isupper()]
    assert len(names) == 61
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_column("A2", [1])
    ws.add_table("A1:A2", px.Table(style=px.TableStyle.None_))
    ws.add_table("C1:C2", px.Table(style=px.TableStyle.Light21))
    path = out()
    wb.save(path)
    sh = reopen(path)["Sheet1"]
    styles = {t.ref: t.tableStyleInfo.name for t in sh.tables.values()}
    assert styles["C1:C2"] == "TableStyleLight21"


def test_no_header_row(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.write_column("A1", [1, 2])
    ws.add_table("A1:A2", px.Table(header_row=False))
    path = out()
    wb.save(path)
    t = next(iter(reopen(path)["Sheet1"].tables.values()))
    assert t.headerRowCount == 0


def test_table_errors():
    ws = px.Workbook().add_worksheet()
    ws.add_table("A1:B3")
    with pytest.raises(px.OverlapError):
        ws.add_table("B2:C4")
    with pytest.raises(px.ParameterError, match="only 2 columns wide"):
        ws.add_table("E1:F3", px.Table(columns=["a", "b", "c"]))
    with pytest.raises(px.ParameterError):
        ws.add_table("E1:E1")  # a header row needs at least one data row
    with pytest.raises(px.ParameterError):
        ws.add_table("H1:I3", px.Table(columns=["Same", "same"]))
    with pytest.raises(px.ParameterError):
        ws.add_table("K1:L3", px.Table(name="bad name"))
    with pytest.raises(px.RangeError):
        ws.add_table("B3:A1")
    with pytest.raises(TypeError):
        ws.add_table("N1:O3", "not a table")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        px.Table(columns=[1, 2])  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        px.Table(columns="AB")
    with pytest.raises(px.ParameterError):
        px.Table(name="")
