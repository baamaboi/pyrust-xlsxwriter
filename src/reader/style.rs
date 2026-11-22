//! Turning a calamine cell `Style` into a `FormatSpec`.
//!
//! calamine resolves theme and indexed colours (including tints) to RGB, so
//! every colour read back is an RGB `Color`. For solid fills the visible
//! colour (OOXML's `fgColor`) is reported as `background_color`, which is
//! how `Format(background_color=...)` writes it; for other patterns
//! `foreground_color` and `background_color` are the pattern's two colours.

use calamine::{
    Border, BorderStyle, Color as CColor, FillPattern, HorizontalAlignment, Style, TextRotation,
    UnderlineStyle, VerticalAlignment,
};
use rust_xlsxwriter::Color as RColor;

use crate::common::color::Color;
use crate::common::format::{
    FormatAlign, FormatBorder, FormatDiagonalBorder, FormatPattern, FormatSpec, FormatUnderline,
};

fn color(c: &CColor) -> Color {
    Color(RColor::RGB(
        ((c.red as u32) << 16) | ((c.green as u32) << 8) | c.blue as u32,
    ))
}

fn border_style(style: BorderStyle) -> Option<FormatBorder> {
    Some(match style {
        BorderStyle::Thin => FormatBorder::Thin,
        BorderStyle::Medium => FormatBorder::Medium,
        BorderStyle::Thick => FormatBorder::Thick,
        BorderStyle::Double => FormatBorder::Double,
        BorderStyle::Hair => FormatBorder::Hair,
        BorderStyle::Dashed => FormatBorder::Dashed,
        BorderStyle::Dotted => FormatBorder::Dotted,
        BorderStyle::MediumDashed => FormatBorder::MediumDashed,
        BorderStyle::DashDot => FormatBorder::DashDot,
        BorderStyle::MediumDashDot => FormatBorder::MediumDashDot,
        BorderStyle::DashDotDot => FormatBorder::DashDotDot,
        BorderStyle::MediumDashDotDot => FormatBorder::MediumDashDotDot,
        BorderStyle::SlantDashDot => FormatBorder::SlantDashDot,
        // `None` and any style added to calamine later.
        _ => return None,
    })
}

fn pattern(p: FillPattern) -> Option<FormatPattern> {
    Some(match p {
        FillPattern::Solid => FormatPattern::Solid,
        FillPattern::DarkGray => FormatPattern::DarkGray,
        FillPattern::MediumGray => FormatPattern::MediumGray,
        FillPattern::LightGray => FormatPattern::LightGray,
        FillPattern::Gray125 => FormatPattern::Gray125,
        FillPattern::Gray0625 => FormatPattern::Gray0625,
        FillPattern::DarkHorizontal => FormatPattern::DarkHorizontal,
        FillPattern::DarkVertical => FormatPattern::DarkVertical,
        FillPattern::DarkDown => FormatPattern::DarkDown,
        FillPattern::DarkUp => FormatPattern::DarkUp,
        FillPattern::DarkGrid => FormatPattern::DarkGrid,
        FillPattern::DarkTrellis => FormatPattern::DarkTrellis,
        FillPattern::LightHorizontal => FormatPattern::LightHorizontal,
        FillPattern::LightVertical => FormatPattern::LightVertical,
        FillPattern::LightDown => FormatPattern::LightDown,
        FillPattern::LightUp => FormatPattern::LightUp,
        FillPattern::LightGrid => FormatPattern::LightGrid,
        FillPattern::LightTrellis => FormatPattern::LightTrellis,
        _ => return None,
    })
}

fn underline(u: UnderlineStyle) -> Option<FormatUnderline> {
    Some(match u {
        UnderlineStyle::Single => FormatUnderline::Single,
        UnderlineStyle::Double => FormatUnderline::Double,
        UnderlineStyle::SingleAccounting => FormatUnderline::SingleAccounting,
        UnderlineStyle::DoubleAccounting => FormatUnderline::DoubleAccounting,
        _ => return None,
    })
}

fn horizontal(h: HorizontalAlignment) -> Option<FormatAlign> {
    Some(match h {
        HorizontalAlignment::Left => FormatAlign::Left,
        HorizontalAlignment::Center => FormatAlign::Center,
        HorizontalAlignment::CenterContinuous => FormatAlign::CenterAcross,
        HorizontalAlignment::Right => FormatAlign::Right,
        HorizontalAlignment::Justify => FormatAlign::Justify,
        HorizontalAlignment::Distributed => FormatAlign::Distributed,
        HorizontalAlignment::Fill => FormatAlign::Fill,
        _ => return None,
    })
}

fn vertical(v: VerticalAlignment) -> Option<FormatAlign> {
    Some(match v {
        VerticalAlignment::Top => FormatAlign::Top,
        VerticalAlignment::Center => FormatAlign::VerticalCenter,
        VerticalAlignment::Justify => FormatAlign::VerticalJustify,
        VerticalAlignment::Distributed => FormatAlign::VerticalDistributed,
        // Bottom is Excel's default.
        _ => return None,
    })
}

/// OOXML rotation (0-90 up, 91-180 down, 255 stacked) to the writer's
/// convention (-90..=90, 270 for stacked text).
fn rotation(r: TextRotation) -> i16 {
    match r {
        TextRotation::Degrees(d @ 0..=90) => d as i16,
        TextRotation::Degrees(d @ 91..=180) => 90 - d as i16,
        TextRotation::Stacked => 270,
        _ => 0,
    }
}

fn side(border: &Border) -> (Option<FormatBorder>, Option<Color>) {
    match border_style(border.style) {
        Some(style) => (Some(style), border.color.as_ref().map(color)),
        None => (None, None),
    }
}

/// Build the `FormatSpec` describing a calamine cell style.
pub fn style_to_spec(style: &Style) -> FormatSpec {
    let mut spec = FormatSpec::default();

    if let Some(font) = &style.font {
        spec.font_name = font.name.clone();
        spec.font_size = font.size;
        spec.bold = font.is_bold();
        spec.italic = font.is_italic();
        spec.strikethrough = font.strikethrough;
        spec.underline = underline(font.underline);
        spec.font_color = font.color.as_ref().map(color);
        spec.font_family = font.family;
    }

    if let Some(fill) = &style.fill
        && let Some(p) = pattern(fill.pattern)
    {
        spec.pattern = Some(p);
        if p == FormatPattern::Solid {
            spec.background_color = fill.foreground_color.as_ref().map(color);
        } else {
            spec.foreground_color = fill.foreground_color.as_ref().map(color);
            spec.background_color = fill.background_color.as_ref().map(color);
        }
    }

    if let Some(borders) = &style.borders {
        (spec.border_top, spec.border_top_color) = side(&borders.top);
        (spec.border_bottom, spec.border_bottom_color) = side(&borders.bottom);
        (spec.border_left, spec.border_left_color) = side(&borders.left);
        (spec.border_right, spec.border_right_color) = side(&borders.right);
        let (up, up_color) = side(&borders.diagonal_up);
        let (down, down_color) = side(&borders.diagonal_down);
        let (kind, style, color) = match (up, down) {
            (Some(s), Some(_)) => (FormatDiagonalBorder::BorderUpDown, Some(s), up_color),
            (Some(s), None) => (FormatDiagonalBorder::BorderUp, Some(s), up_color),
            (None, Some(s)) => (FormatDiagonalBorder::BorderDown, Some(s), down_color),
            (None, None) => (FormatDiagonalBorder::None, None, None),
        };
        if style.is_some() {
            spec.border_diagonal = style;
            spec.border_diagonal_color = color;
            spec.border_diagonal_type = Some(kind);
        }
    }

    if let Some(alignment) = &style.alignment {
        spec.align = horizontal(alignment.horizontal);
        spec.valign = vertical(alignment.vertical);
        spec.rotation = rotation(alignment.text_rotation);
        spec.text_wrap = alignment.wrap_text;
        spec.indent = alignment.indent.unwrap_or(0);
        spec.shrink = alignment.shrink_to_fit;
    }

    if let Some(number_format) = &style.number_format {
        let id = number_format.format_id.unwrap_or(0);
        if number_format.format_code != "General" {
            spec.num_format = Some(number_format.format_code.clone());
        } else if id != 0 && id < 164 {
            // A built-in, locale-dependent format calamine has no code for.
            spec.num_format_index = u8::try_from(id).ok();
        }
    }

    if let Some(protection) = &style.protection {
        spec.locked = protection.locked;
        spec.hidden = protection.hidden;
    }

    spec
}
