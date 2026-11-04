//! Cell styles for [`WriterHandle`](crate::WriterHandle), mirroring the ABI's `xl_cell_style`.

use crate::ffi::{self, XlBorder, XlCellStyle};
use crate::workbook::len_i32;
use crate::Error;

/// The line style of a border edge. Values are Core's `ExcelBorderStyle`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[repr(i32)]
pub enum BorderStyle {
    #[default]
    None = 0,
    Thin = 1,
    Medium = 2,
    Dashed = 3,
    Dotted = 4,
    Thick = 5,
    Double = 6,
    Hair = 7,
    MediumDashed = 8,
    DashDot = 9,
    MediumDashDot = 10,
    DashDotDot = 11,
    MediumDashDotDot = 12,
    SlantDashDot = 13,
}

/// Horizontal alignment. Values are Core's `ExcelHorizontalAlignment`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[repr(i32)]
pub enum HorizontalAlignment {
    /// Text left, numbers right.
    #[default]
    General = 0,
    Left = 1,
    Center = 2,
    Right = 3,
    Fill = 4,
    Justify = 5,
    CenterContinuous = 6,
    Distributed = 7,
}

/// Vertical alignment. Values are Core's `ExcelVerticalAlignment`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[repr(i32)]
pub enum VerticalAlignment {
    #[default]
    Bottom = 0,
    Top = 1,
    Center = 2,
    Justify = 3,
    Distributed = 4,
}

/// One border edge. `color` is `0xRRGGBB`, or `-1` for automatic.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Border {
    pub style: BorderStyle,
    pub color: i32,
}

impl Border {
    #[must_use]
    pub const fn new(style: BorderStyle, color: i32) -> Self {
        Border { style, color }
    }

    fn to_raw(self) -> XlBorder {
        XlBorder {
            style: self.style as i32,
            color: self.color,
        }
    }
}

impl Default for Border {
    fn default() -> Self {
        Border::new(BorderStyle::None, -1)
    }
}

/// A cell style, registered with [`WriterHandle::add_style`](crate::WriterHandle::add_style) or
/// given to a [`Column`](crate::Column) or [`SheetOptions`](crate::SheetOptions). The default is the
/// general style. Colors are `0xRRGGBB`, `-1` for unset; values out of range are rejected by the
/// native library with [`ErrorKind::InvalidArgument`](crate::ErrorKind::InvalidArgument).
///
/// ```
/// use excelreader::{Border, BorderStyle, CellStyle, HorizontalAlignment};
///
/// let header = CellStyle::new()
///     .bold(true)
///     .fill_color(0xDDEBF7)
///     .horizontal_alignment(HorizontalAlignment::Center)
///     .border(Border::new(BorderStyle::Thin, -1));
/// ```
#[derive(Clone, Debug, PartialEq)]
pub struct CellStyle {
    number_format: Option<String>,
    font_name: Option<String>,
    font_flags: i32,
    font_size: f64,
    font_color: i32,
    fill_color: i32,
    horizontal_alignment: HorizontalAlignment,
    vertical_alignment: VerticalAlignment,
    wrap_text: bool,
    indent: i32,
    rotation: i32,
    left: Border,
    right: Border,
    top: Border,
    bottom: Border,
}

impl Default for CellStyle {
    fn default() -> Self {
        CellStyle {
            number_format: None,
            font_name: None,
            font_flags: 0,
            font_size: 0.0,
            font_color: -1,
            fill_color: -1,
            horizontal_alignment: HorizontalAlignment::General,
            vertical_alignment: VerticalAlignment::Bottom,
            wrap_text: false,
            indent: 0,
            rotation: 0,
            left: Border::default(),
            right: Border::default(),
            top: Border::default(),
            bottom: Border::default(),
        }
    }
}

impl CellStyle {
    /// The general style.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// A number-format code such as `"#,##0.00"` or `"yyyy-mm-dd"`. Unset is General.
    #[must_use]
    pub fn number_format(mut self, value: impl Into<String>) -> Self {
        self.number_format = Some(value.into());
        self
    }

    /// The font family, 1-31 characters. Unset is the default font.
    #[must_use]
    pub fn font_name(mut self, value: impl Into<String>) -> Self {
        self.font_name = Some(value.into());
        self
    }

    #[must_use]
    pub fn bold(self, value: bool) -> Self {
        self.font_flag(ffi::XL_FONT_BOLD, value)
    }

    #[must_use]
    pub fn italic(self, value: bool) -> Self {
        self.font_flag(ffi::XL_FONT_ITALIC, value)
    }

    #[must_use]
    pub fn underline(self, value: bool) -> Self {
        self.font_flag(ffi::XL_FONT_UNDERLINE, value)
    }

    #[must_use]
    pub fn strikethrough(self, value: bool) -> Self {
        self.font_flag(ffi::XL_FONT_STRIKETHROUGH, value)
    }

    /// Points, 1-409; `0` is the default size.
    #[must_use]
    pub fn font_size(mut self, value: f64) -> Self {
        self.font_size = value;
        self
    }

    /// `0xRRGGBB`, or `-1` for the default color.
    #[must_use]
    pub fn font_color(mut self, value: i32) -> Self {
        self.font_color = value;
        self
    }

    /// Solid background as `0xRRGGBB`, or `-1` for no fill.
    #[must_use]
    pub fn fill_color(mut self, value: i32) -> Self {
        self.fill_color = value;
        self
    }

    #[must_use]
    pub fn horizontal_alignment(mut self, value: HorizontalAlignment) -> Self {
        self.horizontal_alignment = value;
        self
    }

    #[must_use]
    pub fn vertical_alignment(mut self, value: VerticalAlignment) -> Self {
        self.vertical_alignment = value;
        self
    }

    #[must_use]
    pub fn wrap_text(mut self, value: bool) -> Self {
        self.wrap_text = value;
        self
    }

    /// 0-15.
    #[must_use]
    pub fn indent(mut self, value: i32) -> Self {
        self.indent = value;
        self
    }

    /// Degrees, -90 to 90, or 255 for stacked text.
    #[must_use]
    pub fn rotation(mut self, value: i32) -> Self {
        self.rotation = value;
        self
    }

    /// All four edges.
    #[must_use]
    pub fn border(self, value: Border) -> Self {
        self.left(value).right(value).top(value).bottom(value)
    }

    #[must_use]
    pub fn left(mut self, value: Border) -> Self {
        self.left = value;
        self
    }

    #[must_use]
    pub fn right(mut self, value: Border) -> Self {
        self.right = value;
        self
    }

    #[must_use]
    pub fn top(mut self, value: Border) -> Self {
        self.top = value;
        self
    }

    #[must_use]
    pub fn bottom(mut self, value: Border) -> Self {
        self.bottom = value;
        self
    }

    fn font_flag(mut self, bit: i32, set: bool) -> Self {
        self.font_flags = if set {
            self.font_flags | bit
        } else {
            self.font_flags & !bit
        };
        self
    }

    /// Borrows `self`'s strings, so the result must not outlive it.
    pub(crate) fn to_raw(&self) -> Result<XlCellStyle, Error> {
        let (number_format, number_format_len) =
            text(self.number_format.as_deref(), "number format")?;
        let (font_name, font_name_len) = text(self.font_name.as_deref(), "font name")?;
        Ok(XlCellStyle {
            struct_size: std::mem::size_of::<XlCellStyle>() as i32,
            number_format_len,
            number_format,
            font_name,
            font_name_len,
            font_flags: self.font_flags,
            font_size: self.font_size,
            font_color: self.font_color,
            fill_color: self.fill_color,
            horizontal_alignment: self.horizontal_alignment as i32,
            vertical_alignment: self.vertical_alignment as i32,
            wrap_text: i32::from(self.wrap_text),
            indent: self.indent,
            rotation: self.rotation,
            left: self.left.to_raw(),
            right: self.right.to_raw(),
            top: self.top.to_raw(),
            bottom: self.bottom.to_raw(),
        })
    }
}

fn text(value: Option<&str>, what: &str) -> Result<(*const u8, i32), Error> {
    match value {
        Some(value) => Ok((value.as_ptr(), len_i32(value.len(), what)?)),
        None => Ok((std::ptr::null(), 0)),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_default_style_lowers_with_unset_colors() {
        let raw = CellStyle::new().to_raw().expect("to_raw");
        assert_eq!(raw.struct_size, 104);
        assert!(raw.number_format.is_null());
        assert!(raw.font_name.is_null());
        assert_eq!((raw.font_color, raw.fill_color), (-1, -1));
        for edge in [raw.left, raw.right, raw.top, raw.bottom] {
            assert_eq!((edge.style, edge.color), (0, -1));
        }
    }

    #[test]
    fn every_setter_reaches_the_raw_struct() {
        let style = CellStyle::new()
            .number_format("0.00")
            .font_name("Arial")
            .bold(true)
            .strikethrough(true)
            .italic(true)
            .italic(false)
            .font_size(11.5)
            .font_color(0x123456)
            .fill_color(0)
            .horizontal_alignment(HorizontalAlignment::Distributed)
            .vertical_alignment(VerticalAlignment::Center)
            .wrap_text(true)
            .indent(3)
            .rotation(-90)
            .border(Border::new(BorderStyle::Thin, 0xFF))
            .top(Border::new(BorderStyle::Thick, -1));
        let raw = style.to_raw().expect("to_raw");
        let number_format = unsafe { std::slice::from_raw_parts(raw.number_format, 4) };
        assert_eq!((number_format, raw.number_format_len), (&b"0.00"[..], 4));
        assert_eq!(raw.font_name_len, 5);
        assert_eq!(
            raw.font_flags,
            ffi::XL_FONT_BOLD | ffi::XL_FONT_STRIKETHROUGH
        );
        assert_eq!(raw.font_size, 11.5);
        assert_eq!((raw.font_color, raw.fill_color), (0x123456, 0));
        assert_eq!((raw.horizontal_alignment, raw.vertical_alignment), (7, 2));
        assert_eq!((raw.wrap_text, raw.indent, raw.rotation), (1, 3, -90));
        assert_eq!((raw.left.style, raw.left.color), (1, 0xFF));
        assert_eq!((raw.bottom.style, raw.bottom.color), (1, 0xFF));
        assert_eq!((raw.top.style, raw.top.color), (5, -1));
    }
}
