import pyrust_xlsxwriter
from pyrust_xlsxwriter import Format, Color, FormatAlign, FormatBorder, FormatUnderline


def test_format_properties():
    fmt = Format()

    # Test defaults
    assert fmt.bold is False
    assert fmt.italic is False
    assert fmt.font_name == "Calibri"
    assert fmt.font_size == 11.0
    assert fmt.font_color is None
    assert fmt.num_format == ""
    assert fmt.align is None
    assert fmt.text_wrap is False
    assert fmt.background_color is None
    assert fmt.foreground_color is None
    assert fmt.border is None
    assert fmt.border_color is None
    assert fmt.underline is None
    assert fmt.rotation == 0
    assert fmt.indent == 0
    assert fmt.shrink is False
    assert fmt.unlocked is False
    assert fmt.hidden is False

    # Test setting and getting properties
    fmt.bold = True
    assert fmt.bold is True

    fmt.italic = True
    assert fmt.italic is True

    fmt.font_name = "Arial"
    assert fmt.font_name == "Arial"

    fmt.font_size = 14.5
    assert fmt.font_size == 14.5

    fmt.font_color = Color.Red
    assert fmt.font_color == Color.Red

    fmt.num_format = "$#,##0.00"
    assert fmt.num_format == "$#,##0.00"

    fmt.align = FormatAlign.Center
    assert fmt.align == FormatAlign.Center

    fmt.text_wrap = True
    assert fmt.text_wrap is True

    fmt.background_color = Color.Yellow
    assert fmt.background_color == Color.Yellow

    fmt.foreground_color = Color.Blue
    assert fmt.foreground_color == Color.Blue

    fmt.border = FormatBorder.Thin
    assert fmt.border == FormatBorder.Thin

    fmt.border_color = Color.Black
    assert fmt.border_color == Color.Black

    fmt.underline = FormatUnderline.Single
    assert fmt.underline == FormatUnderline.Single

    fmt.rotation = 45
    assert fmt.rotation == 45

    fmt.indent = 2
    assert fmt.indent == 2

    fmt.shrink = True
    assert fmt.shrink is True

    fmt.unlocked = True
    assert fmt.unlocked is True

    fmt.hidden = True
    assert fmt.hidden is True

    wb = pyrust_xlsxwriter.Workbook()
    ws = wb.add_worksheet()
    ws.write(0, 0, "Formatted Cell", fmt)
    wb.save("test_format_output.xlsx")

    print("ALL Format properties tested & verified successfully!")


if __name__ == "__main__":
    test_format_properties()
