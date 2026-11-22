"""Worksheet.write_records(): header row plus one row per record."""

import collections
import dataclasses
import datetime as dt
import decimal
import types
from collections.abc import Iterator, Mapping
from typing import ClassVar, NamedTuple

import pytest

import pyrust_xlsxwriter as px


def save(wb, out, name="out.xlsx"):
    path = out(name)
    wb.save(path)
    return path


def grid(sheet, rng):
    r1, c1, r2, c2 = rng
    return [[sheet.cell(r + 1, c + 1).value for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]


@dataclasses.dataclass
class Person:
    name: str
    age: int
    kind: ClassVar[str] = "person"

    @property
    def label(self) -> str:
        return f"{self.name} ({self.age})"


class Point(NamedTuple):
    x: int
    y: int

    @property
    def total(self) -> int:
        return self.x + self.y


Pair = collections.namedtuple("Pair", ["left", "right"])


class Plain:
    def __init__(self, a, b):
        self.a = a
        self.b = b
        self._private = "hidden"


class Slotted:
    __slots__ = ("p", "q")

    def __init__(self, p, q):
        self.p = p
        self.q = q


class FrozenMap(Mapping):
    def __init__(self, data):
        self._data = dict(data)

    def __getitem__(self, key):
        return self._data[key]

    def __iter__(self) -> Iterator:
        return iter(self._data)

    def __len__(self) -> int:
        return len(self._data)


def test_dicts(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    records = [
        {"name": "Ann", "age": 31, "city": "Oslo"},
        {"name": "Bob", "city": None},  # missing key and None: blank
        {"age": 5, "name": "Cy", "city": "Rome", "extra": "ignored"},
    ]
    rng = ws.write_records("A1", records)
    assert rng == (0, 0, 3, 2)
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [
        ["name", "age", "city"],
        ["Ann", 31, "Oslo"],
        ["Bob", None, None],
        ["Cy", 5, "Rome"],
    ]
    assert sh["D1"].value is None and sh["D4"].value is None
    assert px.cell_range(*rng) == "A1:C4"


def test_columns_select_and_order(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    records = [{"a": 1, "b": 2, "c": 3}, {"a": 4, "b": 5, "c": 6}]
    rng = ws.write_records(2, "B", records, columns=["c", "a", "zz"])
    assert rng == (2, 1, 4, 3)
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [["c", "a", "zz"], [3, 1, None], [6, 4, None]]


def test_dataclasses(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    people = [Person("Ann", 31), Person("Bob", 27)]
    rng = ws.write_records("A1", people)
    rng2 = ws.write_records("D1", people, columns=["label", "name"])
    book = reopen(save(wb, out))["Sheet1"]
    assert grid(book, rng) == [["name", "age"], ["Ann", 31], ["Bob", 27]]
    assert grid(book, rng2) == [["label", "name"], ["Ann (31)", "Ann"], ["Bob (27)", "Bob"]]


def test_namedtuples(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    rng = ws.write_records("A1", [Point(1, 2), Point(3, 4)])
    rng2 = ws.write_records("D1", [Point(1, 2), Point(3, 4)], columns=["y", "total", "x"])
    rng3 = ws.write_records("H1", (Pair(i, i * 10) for i in range(3)))
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [["x", "y"], [1, 2], [3, 4]]
    assert grid(sh, rng2) == [["y", "total", "x"], [2, 3, 1], [4, 7, 3]]
    assert grid(sh, rng3) == [["left", "right"], [0, 0], [1, 10], [2, 20]]


def test_plain_objects_and_mappings(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    rng = ws.write_records("A1", [Plain(1, "x"), Plain(2, "y")])
    rng2 = ws.write_records("D1", [Slotted(1, 2)], columns=["q", "p"])
    rng3 = ws.write_records("G1", [FrozenMap({"k": 1, "v": 2}), FrozenMap({"k": 3})])
    rng4 = ws.write_records("J1", [types.MappingProxyType({"m": "p"})])
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [["a", "b"], [1, "x"], [2, "y"]]
    assert grid(sh, rng2) == [["q", "p"], [2, 1]]
    assert grid(sh, rng3) == [["k", "v"], [1, 2], [3, None]]
    assert grid(sh, rng4) == [["m"], ["p"]]


def test_mixed_record_types(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    records = [{"name": "Ann", "age": 31}, Person("Bob", 27), {"name": "Cy"}, Person("Di", 40)]
    rng = ws.write_records("A1", records)
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [["name", "age"], ["Ann", 31], ["Bob", 27], ["Cy", None], ["Di", 40]]


def test_value_types(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    record = {
        "date": dt.date(2024, 1, 2),
        "when": dt.datetime(2024, 1, 2, 3, 4),
        "flag": True,
        "dec": decimal.Decimal("1.25"),
        "formula": "=1+1",
        "url": px.Url("https://example.com", text="ex"),
    }
    rng = ws.write_records("A1", [record])
    sh = reopen(save(wb, out))["Sheet1"]
    row = [sh.cell(2, c + 1) for c in range(rng[3] + 1)]
    assert row[0].value == dt.datetime(2024, 1, 2) and row[0].number_format == "yyyy-mm-dd"
    assert row[1].value == dt.datetime(2024, 1, 2, 3, 4)
    assert row[2].value is True
    assert row[3].value == 1.25
    assert row[4].value == "=1+1"
    assert row[5].value == "ex" and row[5].hyperlink.target == "https://example.com"


def test_header_options_and_formats(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    bold = px.Format(bold=True)
    money = px.Format(num_format="0.00")
    italic = px.Format(italic=True)
    records = [{"item": "a", "price": 1.5, "qty": None}, {"item": "b", "price": 2, "qty": 3}]
    rng = ws.write_records("A1", records, header_format=bold, format=italic)
    rng2 = ws.write_records("E1", records, header=False, format={"price": money, "qty": italic, "item": None})
    assert rng == (0, 0, 2, 2) and rng2 == (0, 4, 1, 6)
    sh = reopen(save(wb, out))["Sheet1"]
    assert all(sh.cell(1, c).font.b for c in (1, 2, 3))
    assert sh["A2"].font.i and sh["B3"].font.i and not sh["A1"].font.i
    assert sh["C2"].font.i  # a blank cell still gets the format
    assert grid(sh, rng2) == [["a", 1.5, None], ["b", 2, 3]]
    assert sh["F1"].number_format == "0.00" and not sh["F1"].font.i
    assert sh["G2"].font.i and not sh["E1"].font.i


def test_non_string_keys(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    rng = ws.write_records("A1", [{1: "one", 2.5: "two"}])
    sh = reopen(save(wb, out))["Sheet1"]
    assert grid(sh, rng) == [["1", "2.5"], ["one", "two"]]
    assert sh["A1"].data_type == "s"


def test_empty_records(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    assert ws.write_records("B2", [], columns=["x", "y"]) == (1, 1, 1, 2)
    sh = reopen(save(wb, out))["Sheet1"]
    assert [sh["B2"].value, sh["C2"].value, sh["B3"].value] == ["x", "y", None]
    with pytest.raises(px.ParameterError, match="no records"):
        ws.write_records("A10", [])
    with pytest.raises(px.ParameterError, match="nothing to write"):
        ws.write_records("A10", [], columns=["x"], header=False)
    with pytest.raises(px.ParameterError, match="no columns"):
        ws.write_records("A10", [{}])
    with pytest.raises(px.ParameterError, match="no columns"):
        ws.write_records("A10", [{"a": 1}], columns=[])


def test_record_errors():
    ws = px.Workbook().add_worksheet()
    for bad in ([[1, 2]], [(1, 2)], ["text"], [5], [None]):
        with pytest.raises(TypeError, match="records must be"):
            ws.write_records("A1", bad)
    with pytest.raises(TypeError, match="records must be"):
        ws.write_records("A1", [{"a": 1}, [1]])
    with pytest.raises(TypeError, match="pass columns"):
        ws.write_records("A1", [Slotted(1, 2)])
    with pytest.raises(AttributeError):
        ws.write_records("A1", [Person("Ann", 3)], columns=["name", "missing"])
    with pytest.raises(TypeError, match="attribute names"):
        ws.write_records("A1", [Person("Ann", 3)], columns=[1])
    with pytest.raises(TypeError, match="iterable of records"):
        ws.write_records("A1", 5)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="list of column names"):
        ws.write_records("A1", [{"a": 1}], columns="a")
    with pytest.raises(TypeError, match="as keywords"):
        ws.write_records("A1", [{"a": 1}], ["a"])  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="unexpected keyword"):
        ws.write_records("A1", [{"a": 1}], colums=["a"])  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(TypeError, match="'format' must be a Format"):
        ws.write_records("A1", [{"a": 1}], format="bold")  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError, match="unknown column"):
        ws.write_records("A1", [{"a": 1}], format={"b": px.Format()})
    with pytest.raises(TypeError):
        ws.write_records("A1", [{"a": 1}], header_format="bold")  # pyrefly: ignore[no-matching-overload]


def test_limits():
    ws = px.Workbook().add_worksheet()
    with pytest.raises(px.RangeError, match="columns"):
        ws.write_records(0, 16_383, [{"a": 1, "b": 2}])
    with pytest.raises(px.RangeError):
        ws.write_records(1_048_575, 0, [{"a": 1}])  # header fits, the record doesn't
    assert ws.write_records(1_048_575, 0, [{"a": 1}], header=False) == (1_048_575, 0, 1_048_575, 0)


def test_many_records(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    n = 20_000
    rng = ws.write_records("A1", ({"i": i, "s": f"v{i % 7}", "f": i / 4} for i in range(n)))
    assert rng == (0, 0, n, 2)
    sh = reopen(save(wb, out), read_only=True)["Sheet1"]
    rows = list(sh.iter_rows(values_only=True))
    assert len(rows) == n + 1 and rows[-1] == (n - 1, f"v{(n - 1) % 7}", (n - 1) / 4)
