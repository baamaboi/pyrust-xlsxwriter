"""Images: the Image class, inserted/embedded/fitted images, backgrounds and
header/footer images.

openpyxl drops images when Pillow isn't installed, so files are validated
with ``reopen`` and the drawing parts are checked as XML.
"""

import copy
import io
import re
import struct
import zipfile
import zlib
from pathlib import Path

import pytest

import pyrust_xlsxwriter as px

CORPUS = Path(__file__).parent / "corpus"
EMU = 9525  # EMUs per pixel


def make_png(width: int, height: int, dpi: float | None = None) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(kind + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", crc)

    raw = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    phys = b""
    if dpi is not None:
        per_metre = round(dpi / 0.0254)
        phys = chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + phys
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def make_bmp(width: int = 2, height: int = 2) -> bytes:
    row = (b"\x00\x00\xff" * width).ljust((3 * width + 3) // 4 * 4, b"\x00")
    pixels = row * height
    dib = struct.pack("<IiiHHIIiiII", 40, width, height, 1, 24, 0, len(pixels), 2835, 2835, 0, 0)
    return b"BM" + struct.pack("<IHHI", 54 + len(pixels), 0, 0, 54) + dib + pixels


GIF = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\xff\xff\xff\x00\x00\x00!\xf9\x04\x01\x00\x00\x00\x00"
    b",\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)
PNG = make_png(32, 16)


@pytest.fixture(scope="module")
def jpeg() -> bytes:
    with zipfile.ZipFile(CORPUS / "image19.xlsx") as zf:
        return zf.read("xl/media/image1.jpeg")


def names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        return zf.namelist()


def anchors(xml: str) -> list[str]:
    return re.findall(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", xml)


# ---------------------------------------------------------------------------
# The Image class
# ---------------------------------------------------------------------------


def test_image_from_path(tmp_path):
    path = tmp_path / "logo.png"
    path.write_bytes(PNG)
    for source in (path, str(path)):
        img = px.Image(source)
        assert (img.pixel_width, img.pixel_height) == (32.0, 16.0)
        assert (img.width_dpi, img.height_dpi) == (96.0, 96.0)
        assert img.image_type == "png"
        assert img.data == PNG
        assert (img.width, img.height) == (32.0, 16.0)
        assert (img.scale_width, img.scale_height) == (1.0, 1.0)


def test_image_from_bytes_variants(jpeg):
    for data in (PNG, bytearray(PNG), memoryview(PNG), io.BytesIO(PNG)):
        assert px.Image.from_bytes(data).pixel_width == 32
    assert px.Image.from_buffer(PNG) == px.Image.from_bytes(PNG)
    fh = io.BufferedReader(io.BytesIO(PNG))
    assert px.Image.from_bytes(fh).pixel_height == 16
    with pytest.raises(TypeError, match="Image\\(path\\)"):
        px.Image.from_bytes("logo.png")  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError):
        px.Image.from_bytes(12)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="binary mode"):
        px.Image.from_bytes(io.StringIO("text"))  # pyrefly: ignore[bad-argument-type]
    img = px.Image.from_bytes(jpeg)
    assert (img.image_type, img.pixel_width, img.pixel_height) == ("jpeg", 640, 480)


@pytest.mark.parametrize(
    ("data", "kind", "size"),
    [(GIF, "gif", (1, 1)), (make_bmp(2, 3), "bmp", (2, 3)), (make_png(5, 7), "png", (5, 7))],
)
def test_image_types(data, kind, size):
    img = px.Image.from_bytes(data)
    assert img.image_type == kind
    assert (img.pixel_width, img.pixel_height) == size


def test_image_dpi():
    img = px.Image.from_bytes(make_png(10, 10, dpi=144))
    assert img.width_dpi == pytest.approx(144, abs=0.01)
    assert img.height_dpi == pytest.approx(144, abs=0.01)


def test_bad_image_data():
    for data in (b"", b"not an image at all, definitely not one", b"\x89PNG" + b"\x00" * 40):
        with pytest.raises(px.ImageError):
            px.Image.from_bytes(data)
    # ImageError is a ParameterError / ValueError.
    with pytest.raises(ValueError, match="unknown image type"):
        px.Image.from_bytes(b"x" * 64)
    # A PNG with zero dimensions.
    with pytest.raises(px.ImageError, match="dimensions"):
        px.Image.from_bytes(make_png(0, 0))


def test_missing_file(tmp_path):
    with pytest.raises(px.FileError) as info:
        px.Image(tmp_path / "missing.png")
    assert isinstance(info.value, OSError)
    assert "missing.png" in str(info.value)
    bad = tmp_path / "bad.png"
    bad.write_bytes(b"x" * 100)
    with pytest.raises(px.ImageError, match="bad.png"):
        px.Image(bad)


def test_size_and_scale_properties():
    img = px.Image.from_bytes(PNG)
    img.width = 64
    assert img.scale_width == 2.0 and img.width == 64.0
    img.scale_height = 0.5
    assert img.height == 8.0
    img.height = 48
    assert img.scale_height == 3.0
    for name in ("width", "height", "scale_width", "scale_height"):
        for bad in (0, -1, float("nan"), float("inf")):
            with pytest.raises(px.ParameterError):
                setattr(img, name, bad)
    with pytest.raises(px.ParameterError, match="limit"):
        img.scale_width = 1e6
    assert img.scale_width == 2.0  # unchanged after a failed set


def test_scale_to_size():
    img = px.Image.from_bytes(PNG)  # 32x16
    img.scale_to_size(64, 64)
    assert (img.width, img.height) == (64, 64)
    img.scale_to_size(64, 64, keep_aspect_ratio=True)
    assert (img.width, img.height) == (64, 32)
    img.scale_to_size(10, 100, True)
    assert img.scale_width == img.scale_height == pytest.approx(10 / 32)
    # DPI is taken into account, as in rust_xlsxwriter.
    hi = px.Image.from_bytes(make_png(100, 100, dpi=192))
    hi.scale_to_size(50, 50)
    assert hi.scale_width == pytest.approx(1.0, abs=0.001)
    with pytest.raises(px.ParameterError):
        img.scale_to_size(0, 10)


def test_kwargs_and_other_properties():
    img = px.Image.from_bytes(
        PNG,
        width=64,
        scale_height=2,
        alt_text="Logo",
        decorative=True,
        object_movement=px.ObjectMovement.DontMoveOrSizeWithCells,
        url="https://example.com",
    )
    assert img.scale_width == 2 and img.scale_height == 2
    assert img.alt_text == "Logo"
    assert img.decorative is True
    assert img.object_movement == px.ObjectMovement.DontMoveOrSizeWithCells
    assert isinstance(img.url, px.Url) and img.url.link == "https://example.com"
    img.url = px.Url("internal:Sheet1!A1", tip="Go")
    assert isinstance(img.url, px.Url) and img.url.tip == "Go"
    img.url = None
    assert img.url is None
    with pytest.raises(TypeError, match="unexpected keyword"):
        px.Image.from_bytes(PNG, colour="red")  # pyrefly: ignore[unexpected-keyword]


def test_validation():
    img = px.Image.from_bytes(PNG)
    with pytest.raises(px.ParameterError, match="255"):
        img.alt_text = "x" * 256
    img.alt_text = "x" * 255
    with pytest.raises(px.ParameterError):
        img.url = "not a url"
    with pytest.raises(TypeError):
        img.url = 5  # pyrefly: ignore[bad-argument-type]


def test_copy_eq_repr():
    img = px.Image.from_bytes(PNG, alt_text="a")
    other = img.copy()
    assert other == img and other is not img
    other.alt_text = "b"
    assert other != img and img.alt_text == "a"
    assert copy.copy(img) == img and copy.deepcopy(img) == img
    assert px.Image.from_bytes(PNG) != px.Image.from_bytes(make_png(32, 17))
    assert img != "image"
    assert repr(img) == "<Image png 32x16 scale=(1, 1)>"


# ---------------------------------------------------------------------------
# Worksheet methods
# ---------------------------------------------------------------------------


def test_insert_image(out, reopen, part, tmp_path):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    img = px.Image.from_bytes(PNG, alt_text="Logo", url="https://example.com")
    ws.insert_image("B2", img)
    ws.insert_image(4, 1, img, x_offset=5, y_offset=7)
    ws.insert_image(8, "B", img, 3)
    ws.insert_image_with_offset("D2", img, 10, 2)
    ws.insert_image_with_offset(9, 3, image=img, x_offset=1, y_offset=1)
    path_img = tmp_path / "pic.png"
    path_img.write_bytes(make_png(8, 8))
    ws.insert_image("F2", path_img)  # a path works too
    ws.insert_image("F4", str(path_img))
    path = out()
    wb.save(path)
    reopen(path)
    xml = part(path, "xl/drawings/drawing1.xml")
    items = anchors(xml)
    assert len(items) == 7
    first = items[0]
    assert "<xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>" in first
    assert f'<a:ext cx="{32 * EMU}" cy="{16 * EMU}"/>' in first
    assert 'descr="Logo"' in first and "a:hlinkClick" in first
    assert f"<xdr:colOff>{5 * EMU}</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>{7 * EMU}</xdr:rowOff>" in xml
    assert f"<xdr:col>3</xdr:col><xdr:colOff>{10 * EMU}</xdr:colOff>" in xml
    rels = part(path, "xl/drawings/_rels/drawing1.xml.rels")
    assert 'Target="https://example.com"' in rels
    media = [n for n in names(path) if n.startswith("xl/media/")]
    assert len(media) == 2  # identical images are stored once


def test_insert_image_errors():
    ws = px.Workbook().add_worksheet()
    img = px.Image.from_bytes(PNG)
    with pytest.raises(TypeError, match="missing required argument: 'x_offset'"):
        ws.insert_image_with_offset("A1", img)  # pyrefly: ignore[no-matching-overload]
    with pytest.raises(px.ParameterError, match="x_offset"):
        ws.insert_image("A1", img, x_offset=-1)
    with pytest.raises(px.ParameterError, match="y_offset"):
        ws.insert_image("A1", img, y_offset=5_000_000)
    with pytest.raises(TypeError, match="Image.from_bytes"):
        ws.insert_image("A1", PNG)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(TypeError, match="expected an Image"):
        ws.insert_image("A1", 3.5)  # pyrefly: ignore[bad-argument-type]
    with pytest.raises(px.FileError):
        ws.insert_image("A1", "/no/such/image.png")
    with pytest.raises(px.RangeError):
        ws.insert_image(1_048_576, 0, img)
    with pytest.raises(px.ParameterError):
        ws.insert_image("A0", img)


def test_object_movement_and_decorative(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    expected = {
        px.ObjectMovement.MoveAndSizeWithCells: None,
        px.ObjectMovement.MoveButDontSizeWithCells: "oneCell",
        px.ObjectMovement.DontMoveOrSizeWithCells: "absolute",
        px.ObjectMovement.MoveAndSizeWithCellsAfter: None,
    }
    for row, movement in enumerate(expected):
        ws.insert_image(row * 3, 0, px.Image.from_bytes(PNG, object_movement=movement))
    ws.insert_image("A20", px.Image.from_bytes(PNG, decorative=True))
    path = out()
    wb.save(path)
    reopen(path)
    items = anchors(part(path, "xl/drawings/drawing1.xml"))
    for item, edit_as in zip(items, expected.values()):
        if edit_as is None:
            assert "editAs" not in item.split(">")[0]
        else:
            assert f'editAs="{edit_as}"' in item
    assert "adec:decorative" in items[-1]


def test_all_image_types(out, reopen, part, jpeg):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    for row, data in enumerate((PNG, jpeg, GIF, make_bmp())):
        ws.insert_image(row * 10, 0, px.Image.from_bytes(data))
    path = out()
    wb.save(path)
    reopen(path)
    media = sorted(n.rsplit(".", 1)[1] for n in names(path) if n.startswith("xl/media/"))
    assert media == ["bmp", "gif", "jpeg", "png"]
    types = part(path, "[Content_Types].xml")
    for ext in ("png", "jpeg", "gif", "bmp"):
        assert f'Extension="{ext}"' in types


def test_embed_image(out, reopen, part, jpeg):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.embed_image("B2", px.Image.from_bytes(PNG, alt_text="emb"))
    ws.embed_image(3, 1, px.Image.from_bytes(jpeg), px.Format(border=px.FormatBorder.Thin))
    ws.embed_image(5, 1, px.Image.from_bytes(PNG), format=None)
    path = out()
    wb.save(path)
    reopen(path)
    sheet = part(path, "xl/worksheets/sheet1.xml")
    assert re.search(r'<c r="B2"[^>]* t="e" vm="1"><v>#VALUE!</v></c>', sheet)
    assert re.search(r'<c r="B4" s="1" t="e" vm="2">', sheet)
    assert re.search(r'<c r="B6"[^>]* vm="1">', sheet)  # same image, same rich value
    parts = names(path)
    assert "xl/metadata.xml" in parts and "xl/richData/rdrichvalue.xml" in parts
    assert "emb" in part(path, "xl/richData/rdrichvalue.xml")
    assert "xl/drawings/drawing1.xml" not in parts


def test_insert_image_fit_to_cell(out, reopen, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_column_width_pixels("B", 100)
    ws.set_row_height_pixels(1, 40)
    img = px.Image.from_bytes(PNG)  # 32x16
    ws.insert_image_fit_to_cell("B2", img)  # keeps aspect ratio: 80x40
    ws.insert_image_fit_to_cell(1, 2, img, keep_aspect_ratio=False)  # 64x40 (default column)
    ws.insert_image_fit_to_cell("B4", img, centered=True)  # 64x20 row: 40x20
    path = out()
    wb.save(path)
    reopen(path)
    items = anchors(part(path, "xl/drawings/drawing1.xml"))
    assert f'cx="{80 * EMU}" cy="{40 * EMU}"' in items[0]
    assert f'cx="{64 * EMU}" cy="{40 * EMU}"' in items[1]
    assert f'cx="{40 * EMU}" cy="{20 * EMU}"' in items[2]
    assert f"<xdr:colOff>{30 * EMU}</xdr:colOff>" in items[2]  # centered: (100 - 40) / 2


def test_fit_to_hidden_cell_centered(out, reopen):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.set_column_hidden("C")
    tall = px.Image.from_bytes(make_png(32, 32))  # taller than the row
    with pytest.raises(px.ParameterError, match="hidden"):
        ws.insert_image_fit_to_cell("C3", tall, centered=True)
    ws.set_row_hidden(9)
    with pytest.raises(px.ParameterError, match="hidden"):
        ws.insert_image_fit_to_cell("A10", tall, centered=True)
    # Not centered: allowed (the image keeps its size).
    ws.insert_image_fit_to_cell("C3", tall)
    path = out()
    wb.save(path)
    reopen(path)


def test_background_image(out, reopen, part, jpeg):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.insert_background_image(px.Image.from_bytes(jpeg))
    path = out()
    wb.save(path)
    reopen(path)
    sheet = part(path, "xl/worksheets/sheet1.xml")
    picture = re.search(r'<picture r:id="(rId\d+)"/>', sheet)
    assert picture
    rid = picture.group(1)
    rels = part(path, "xl/worksheets/_rels/sheet1.xml.rels")
    assert re.search(rf'Id="{rid}" Type="[^"]*/image" Target="../media/image1.jpeg"', rels)


def test_header_and_footer_images(out, reopen, part, tmp_path):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    ws.header = "&L&[Picture]&C&G"
    ws.footer = "&R&[Picture]"
    logo = tmp_path / "watermark.png"
    logo.write_bytes(make_png(40, 20))
    ws.set_header_image(px.Image(logo), px.HeaderImagePosition.Left)
    ws.set_header_image(px.Image.from_bytes(PNG), px.HeaderImagePosition.Center)
    ws.set_footer_image(str(logo), px.HeaderImagePosition.Right)
    path = out()
    wb.save(path)
    reopen(path)
    sheet = part(path, "xl/worksheets/sheet1.xml")
    assert "<oddHeader>&amp;L&amp;G&amp;C&amp;G</oddHeader>" in sheet
    assert "<oddFooter>&amp;R&amp;G</oddFooter>" in sheet
    assert "<legacyDrawingHF " in sheet
    vml = part(path, "xl/drawings/vmlDrawing1.vml")
    assert re.findall(r'<v:shape id="(\w+)"', vml) == ["LH", "CH", "RF"]
    assert 'o:title="watermark"' in vml and 'o:title="image"' in vml
    assert "width:30pt;height:15pt" in vml  # 40x20 px


def test_header_image_needs_placeholder():
    ws = px.Workbook().add_worksheet()
    img = px.Image.from_bytes(PNG)
    with pytest.raises(px.ParameterError, match=r'ws.header = "&C&\[Picture\]"'):
        ws.set_header_image(img, px.HeaderImagePosition.Center)
    ws.header = "&L&[Picture]"
    with pytest.raises(px.ParameterError, match="&R"):
        ws.set_header_image(img, px.HeaderImagePosition.Right)
    ws.set_header_image(img, px.HeaderImagePosition.Left)
    with pytest.raises(px.ParameterError, match="footer"):
        ws.set_footer_image(img, px.HeaderImagePosition.Left)


def test_header_image_position_enum():
    assert [
        int(p) for p in (px.HeaderImagePosition.Left, px.HeaderImagePosition.Center, px.HeaderImagePosition.Right)
    ] == [0, 1, 2]


def test_image_changes_after_insert_do_not_affect_sheet(out, part):
    wb = px.Workbook()
    ws = wb.add_worksheet()
    img = px.Image.from_bytes(PNG)
    ws.insert_image("A1", img)
    img.width = 320  # applies to later insertions only
    ws.insert_image("A10", img)
    path = out()
    wb.save(path)
    items = anchors(part(path, "xl/drawings/drawing1.xml"))
    assert f'cx="{32 * EMU}"' in items[0] and f'cx="{320 * EMU}"' in items[1]
