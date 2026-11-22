"""Color, Format, cell-reference utilities and the exception hierarchy."""

import copy

import pytest

import pyrust_xlsxwriter as px


class TestColor:
    def test_named_and_parsed(self):
        assert px.Color.Red.hex == "#FF0000"
        assert px.Color("red") == px.Color.Red
        assert px.Color("#4472c4").hex == "#4472C4"
        assert px.Color("4472C4") == px.Color(0x4472C4)
        assert px.Color((68, 114, 196)) == px.Color.rgb(68, 114, 196) == px.Color.rgb(0x4472C4)

    def test_theme(self):
        c = px.Color.theme(3, 2)
        assert c.kind == "theme"
        assert c.theme_index == (3, 2)
        assert c.hex is None

    @pytest.mark.parametrize("bad", ["#12345", "notacolor", (1, 2), (256, 0, 0), -1, 0x1000000])
    def test_invalid(self, bad):
        with pytest.raises((ValueError, TypeError)):
            px.Color(bad)

    def test_hashable(self):
        assert len({px.Color.Red, px.Color("#FF0000"), px.Color.Blue}) == 3  # named Red != RGB red
        assert repr(px.Color.theme(1, 0)) == "Color.theme(1, 0)"


class TestFormat:
    def test_kwargs_and_properties(self):
        f = px.Format(bold=True, font_color="#9C0006", border=px.FormatBorder.Thin)
        assert f.bold is True
        assert f.font_color == px.Color("#9C0006")
        assert f.border == px.FormatBorder.Thin
        assert f.border_left == px.FormatBorder.Thin

    def test_unknown_kwarg(self):
        with pytest.raises(TypeError, match="unexpected keyword"):
            px.Format(blod=True)  # pyrefly: ignore[unexpected-keyword]

    def test_align_routes_vertical(self):
        f = px.Format(align=px.FormatAlign.Center)
        f.align = px.FormatAlign.VerticalCenter
        assert f.align == px.FormatAlign.Center
        assert f.valign == px.FormatAlign.VerticalCenter
        with pytest.raises(ValueError):
            f.valign = px.FormatAlign.Left

    def test_set_copy_merge_eq(self):
        a = px.Format().set(bold=True)
        b = copy.copy(a)
        assert a == b and a is not b
        b.italic = True
        assert a != b
        merged = px.Format(font_size=14).merge(px.Format(bold=True))
        assert merged.font_size == 14 and merged.bold
        assert px.Format(bold=True).to_dict() == {"bold": True}
        assert "bold=True" in repr(a)

    def test_unlocked_alias(self):
        f = px.Format(unlocked=True)
        assert f.locked is False


class TestCellRefs:
    def test_round_trips(self):
        assert px.column_number_to_name(27) == "AB"
        assert px.column_name_to_number("AB") == 27
        assert px.row_col_to_cell(0, 0) == "A1"
        assert px.row_col_to_cell(2, 1, True, True) == "$B$3"
        assert px.cell_to_row_col("$B$3") == (2, 1)
        assert px.cell_range(0, 0, 3, 2) == "A1:C4"
        assert px.range_to_row_col("Sheet1!A1:C4") == (0, 0, 3, 2)
        assert px.quote_sheet_name("My Sheet") == "'My Sheet'"

    def test_invalid(self):
        with pytest.raises(px.ParameterError):
            px.cell_to_row_col("A0")


def test_exception_hierarchy():
    assert issubclass(px.ParameterError, px.XlsxError)
    assert issubclass(px.ParameterError, ValueError)
    assert issubclass(px.SheetNameError, px.ParameterError)
    assert issubclass(px.RangeError, IndexError)
    assert issubclass(px.FileError, OSError)
    assert issubclass(px.ReadError, px.XlsxError)
    assert issubclass(px.EditError, px.XlsxError)
