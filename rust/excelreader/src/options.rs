//! Builder for [`XlOpenOptions`], mirroring the Python binding's `OpenOptions` and the C++
//! wrapper's `xl::OpenOptions`.
//!
//! The ABI splits "unset" across two conventions: `0` for a numeric field, `XL_OPT_DEFAULT` for a
//! boolean-shaped one. This type presents the single `None` convention the other bindings expose and
//! does the split in [`OpenOptions::to_raw`].
//!
//! Values are passed through unvalidated on purpose: the native side owns the real bounds and
//! reports a rejection through `xl_last_error`, so checking them here too would give those bounds a
//! second place to drift from.

use crate::{XL_OPT_DEFAULT, XL_OPT_FALSE, XL_OPT_TRUE, XlOpenOptions, XlWriteOptions};

/// Options for [`Workbook::open_with`](crate::workbook::Workbook::open_with) and
/// [`Workbook::open_memory`](crate::workbook::Workbook::open_memory). Every field is `None` by
/// default, meaning "use the library default"; set only the ones you want to override.
///
/// ```no_run
/// use excelreader::OpenOptions;
///
/// let options = OpenOptions::new()
///     .prefetch_decompression(true)
///     .max_zip_entries(1024);
/// ```
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct OpenOptions {
    // CSV only (format == XL_FORMAT_CSV); ignored for every other format.
    pub csv_sniff_dialect: Option<bool>,
    pub csv_delimiter: Option<i32>,
    pub csv_quote: Option<i32>,
    pub csv_detect_bom: Option<bool>,
    pub csv_max_cell_bytes: Option<i32>,
    pub csv_intern_strings: Option<bool>,

    // XLS/XLSX/XLSB only; ignored for CSV.
    pub max_total_decompressed_bytes: Option<i64>,
    pub max_cell_bytes: Option<i32>,
    pub max_shared_string_bytes: Option<i64>,
    pub max_zip_entries: Option<i32>,
    pub prefetch_decompression: Option<bool>,
    pub intern_strings: Option<bool>,
}

/// Generates a consuming builder setter per field, so callers can chain overrides.
macro_rules! setters {
    ($($(#[$doc:meta])* $name:ident: $ty:ty),* $(,)?) => {
        $(
            $(#[$doc])*
            #[must_use]
            pub fn $name(mut self, value: $ty) -> Self {
                self.$name = Some(value);
                self
            }
        )*
    };
}

impl OpenOptions {
    /// Every field unset - identical to passing no options at all.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    setters! {
        /// CSV only. Infers delimiter/quote/encoding from a leading sample before opening, which
        /// makes `csv_delimiter`/`csv_quote` ignored.
        csv_sniff_dialect: bool,
        /// CSV only. Byte value 1-255; the library default is `,`.
        csv_delimiter: i32,
        /// CSV only. Byte value 1-255; the library default is `"`.
        csv_quote: i32,
        /// CSV only. Library default is `true`.
        csv_detect_bom: bool,
        /// CSV only. Library default is 32 MiB.
        csv_max_cell_bytes: i32,
        /// CSV only. Library default is `false`.
        csv_intern_strings: bool,
        /// XLS/XLSX/XLSB only. Library default is 512 MiB.
        max_total_decompressed_bytes: i64,
        /// XLS/XLSX/XLSB only. Library default is 32 MiB.
        max_cell_bytes: i32,
        /// XLS/XLSX/XLSB only. Library default is 128 MiB.
        max_shared_string_bytes: i64,
        /// XLS/XLSX/XLSB only. Library default is 65536.
        max_zip_entries: i32,
        /// XLS/XLSX/XLSB only. Library default is `false`.
        prefetch_decompression: bool,
        /// XLS/XLSX/XLSB only. Library default is `false`.
        intern_strings: bool,
    }

    /// Lowers this into the raw ABI struct, with `struct_size` filled in.
    #[must_use]
    pub fn to_raw(&self) -> XlOpenOptions {
        XlOpenOptions {
            struct_size: std::mem::size_of::<XlOpenOptions>() as i32,
            csv_sniff_dialect: opt_state(self.csv_sniff_dialect),
            csv_delimiter: opt_number(self.csv_delimiter),
            csv_quote: opt_number(self.csv_quote),
            csv_detect_bom: opt_state(self.csv_detect_bom),
            csv_max_cell_bytes: opt_number(self.csv_max_cell_bytes),
            csv_intern_strings: opt_state(self.csv_intern_strings),
            max_total_decompressed_bytes: opt_number(self.max_total_decompressed_bytes),
            max_cell_bytes: opt_number(self.max_cell_bytes),
            max_shared_string_bytes: opt_number(self.max_shared_string_bytes),
            max_zip_entries: opt_number(self.max_zip_entries),
            prefetch_decompression: opt_state(self.prefetch_decompression),
            intern_strings: opt_state(self.intern_strings),
        }
    }
}

fn opt_state(value: Option<bool>) -> i32 {
    match value {
        None => XL_OPT_DEFAULT,
        Some(true) => XL_OPT_TRUE,
        Some(false) => XL_OPT_FALSE,
    }
}

/// Explicit `None` test rather than `unwrap_or(0)` over a falsy check: the ABI spends the value 0 on
/// "use the default", and no field here has a meaningful 0 (a zero delimiter, or a zero-byte cell
/// limit, is not a setting).
fn opt_number<T: Default>(value: Option<T>) -> T {
    value.unwrap_or_default()
}

/// Options for [`write_columns`](crate::writer::write_columns) and
/// [`write_sheet`](crate::writer::write_sheet). Every field is `None` by default, meaning "use the
/// library default"; set only the ones you want to override.
///
/// ```no_run
/// use excelreader::WriteOptions;
///
/// let options = WriteOptions::new().sheet_name("Dados").use_shared_strings(true);
/// ```
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct WriteOptions {
    /// Sheet name. `None` = `"Sheet1"`. Ignored for CSV. Excel's rules (1-31 characters, none of
    /// `: \ / ? * [ ]`) are enforced natively and reported through `xl_last_error`.
    pub sheet_name: Option<String>,

    // CSV only; ignored for every other format. Byte value 1-255.
    pub csv_delimiter: Option<i32>,
    pub csv_quote: Option<i32>,

    /// XLS/XLSB only; ignored for XLSX and CSV. Library default is `false`.
    pub date1904: Option<bool>,
    /// XLSX/XLSB only; ignored for XLS and CSV. Shrinks files with many repeated strings, at the
    /// cost of a string table. Library default is `false`.
    pub use_shared_strings: Option<bool>,
}

impl WriteOptions {
    /// Every field unset - identical to passing no options at all.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Written by hand rather than generated by `setters!`: this is the one field that is not a
    /// plain `Copy` scalar.
    #[must_use]
    pub fn sheet_name(mut self, value: impl Into<String>) -> Self {
        self.sheet_name = Some(value.into());
        self
    }

    setters! {
        /// CSV only. Byte value 1-255; the library default is `,`.
        csv_delimiter: i32,
        /// CSV only. Byte value 1-255; the library default is `"`.
        csv_quote: i32,
        /// XLS/XLSB only. Library default is `false`.
        date1904: bool,
        /// XLSX/XLSB only. Library default is `false`.
        use_shared_strings: bool,
    }

    /// Lowers this into the raw ABI struct, with `struct_size` filled in.
    ///
    /// `pub(crate)`, unlike [`OpenOptions::to_raw`], which is public: the struct this returns holds
    /// a raw pointer INTO `self.sheet_name`, so handing it to a caller who can outlive `self` is a
    /// use-after-free waiting to happen. `write_columns` takes `Option<&WriteOptions>` and does the
    /// lowering itself, where the borrow provably covers the FFI call.
    pub(crate) fn to_raw(&self) -> XlWriteOptions {
        XlWriteOptions {
            struct_size: std::mem::size_of::<XlWriteOptions>() as i32,
            sheet_name_len: self.sheet_name.as_ref().map_or(0, |n| n.len() as i32),
            sheet_name: self
                .sheet_name
                .as_ref()
                .map_or(std::ptr::null(), |n| n.as_ptr()),
            csv_delimiter: opt_number(self.csv_delimiter),
            csv_quote: opt_number(self.csv_quote),
            date1904: opt_state(self.date1904),
            use_shared_strings: opt_state(self.use_shared_strings),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unset_fields_lower_to_the_abi_defaults() {
        let raw = OpenOptions::new().to_raw();
        assert_eq!(raw.struct_size, std::mem::size_of::<XlOpenOptions>() as i32);
        assert_eq!(raw.csv_sniff_dialect, XL_OPT_DEFAULT);
        assert_eq!(raw.prefetch_decompression, XL_OPT_DEFAULT);
        assert_eq!(raw.csv_delimiter, 0);
        assert_eq!(raw.max_total_decompressed_bytes, 0);
    }

    #[test]
    fn booleans_lower_to_distinct_true_and_false_states() {
        let raw = OpenOptions::new()
            .prefetch_decompression(true)
            .intern_strings(false)
            .to_raw();
        assert_eq!(raw.prefetch_decompression, XL_OPT_TRUE);
        assert_eq!(raw.intern_strings, XL_OPT_FALSE);
        // Still "unset", and distinguishable from the explicit `false` above.
        assert_eq!(raw.csv_detect_bom, XL_OPT_DEFAULT);
    }

    #[test]
    fn numbers_pass_through_unchanged() {
        let raw = OpenOptions::new()
            .csv_delimiter(b';' as i32)
            .max_zip_entries(1024)
            .max_shared_string_bytes(1 << 20)
            .to_raw();
        assert_eq!(raw.csv_delimiter, b';' as i32);
        assert_eq!(raw.max_zip_entries, 1024);
        assert_eq!(raw.max_shared_string_bytes, 1 << 20);
    }

    #[test]
    fn write_options_lower_to_the_abi_defaults() {
        let raw = WriteOptions::new().to_raw();
        assert_eq!(
            raw.struct_size,
            std::mem::size_of::<crate::XlWriteOptions>() as i32
        );
        assert!(raw.sheet_name.is_null(), "an unset sheet name means Sheet1");
        assert_eq!(raw.sheet_name_len, 0);
        assert_eq!(raw.csv_delimiter, 0);
        assert_eq!(raw.date1904, XL_OPT_DEFAULT);
    }

    #[test]
    fn write_options_lower_every_set_field() {
        let options = WriteOptions::new()
            .sheet_name("Dados")
            .csv_delimiter(b';' as i32)
            .date1904(false)
            .use_shared_strings(true);
        let raw = options.to_raw();
        assert_eq!(raw.sheet_name_len, 5);
        assert!(!raw.sheet_name.is_null());
        assert_eq!(raw.csv_delimiter, b';' as i32);
        assert_eq!(raw.date1904, XL_OPT_FALSE);
        assert_eq!(raw.use_shared_strings, XL_OPT_TRUE);
    }

    /// The C struct's x64 layout is pinned by XL_STATIC_ASSERT in
    /// tests/ExcelReader.NativeSmoke/smoke.c (lines 66-73). A repr(C) struct that disagrees would
    /// hand the native side garbage where it expects a length and a pointer.
    #[test]
    fn write_options_match_the_c_struct_size() {
        assert_eq!(std::mem::size_of::<crate::XlWriteOptions>(), 32);
    }
}
