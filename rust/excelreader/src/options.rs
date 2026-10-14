//! Builders for the ABI's open, write and CSV-aggregation option structs, mirroring the Python
//! binding's `OpenOptions` and the C++ wrapper's `xl::OpenOptions`.
//!
//! The ABI splits "unset" across two conventions: `0` for a numeric field, `XL_OPT_DEFAULT` for a
//! boolean-shaped one. These types present the single `None` convention the other bindings expose
//! and do the split when lowering.
//!
//! Values are passed through unvalidated on purpose: the native side owns the real bounds and
//! reports a rejection through `xl_last_error`, so checking them here too would give those bounds a
//! second place to drift from.

use std::path::Path;

use crate::ffi::{
    self, XlCsvParallelOptions, XlOpenOptions, XlWriteOptions, XL_OPT_DEFAULT, XL_OPT_FALSE,
    XL_OPT_TRUE,
};
use crate::{Error, ErrorKind};

/// A workbook file format.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Format {
    /// Opening: sniff the format from the file's signature bytes. Sniffing does not detect CSV.
    /// Keep `Auto` for a password-protected workbook: an explicit `Xlsx`/`Xlsb` reads its
    /// encryption wrapper as a plain ZIP and fails. Writing: infer the format from the path's
    /// extension.
    #[default]
    Auto,
    Xls,
    Xlsx,
    Xlsb,
    Csv,
}

impl Format {
    pub(crate) fn raw(self) -> i32 {
        match self {
            Format::Auto => ffi::XL_FORMAT_AUTO,
            Format::Xls => ffi::XL_FORMAT_XLS,
            Format::Xlsx => ffi::XL_FORMAT_XLSX,
            Format::Xlsb => ffi::XL_FORMAT_XLSB,
            Format::Csv => ffi::XL_FORMAT_CSV,
        }
    }

    /// Infers a format from a path's extension, case-insensitively. `Auto` when the extension is
    /// absent or unrecognized - which the native write then rejects, since a file being created has
    /// no signature bytes to sniff.
    fn from_path(path: &Path) -> Format {
        let extension = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or_default();
        match extension.to_ascii_lowercase().as_str() {
            "xlsx" => Format::Xlsx,
            "xlsb" => Format::Xlsb,
            "xls" => Format::Xls,
            "csv" => Format::Csv,
            _ => Format::Auto,
        }
    }
}

/// How to open a workbook, in the style of [`std::fs::OpenOptions`]: build one, then call
/// [`open`](Self::open), [`open_bytes`](Self::open_bytes), [`open_source`](Self::open_source) or
/// [`open_reader`](Self::open_reader). Every setting is unset by default, meaning "use the library
/// default"; set only the ones you want to override.
///
/// ```no_run
/// use excelreader::{Format, OpenOptions};
///
/// let workbook = OpenOptions::new()
///     .format(Format::Csv)
///     .csv_delimiter(b';' as i32)
///     .open("data.csv")?;
/// # Ok::<(), excelreader::Error>(())
/// ```
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OpenOptions {
    format: Format,

    csv_sniff_dialect: Option<bool>,
    csv_delimiter: Option<i32>,
    csv_quote: Option<i32>,
    csv_detect_bom: Option<bool>,

    max_total_decompressed_bytes: Option<i64>,
    max_cell_bytes: Option<i32>,
    max_shared_string_bytes: Option<i64>,
    max_zip_entries: Option<i32>,
    prefetch_decompression: Option<bool>,
    intern_strings: Option<bool>,
    verify_zip_checksums: Option<bool>,

    source_block_size: Option<i32>,
    source_cache_bytes: Option<i64>,
    max_buffered_bytes: Option<i64>,

    password: Option<String>,
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
    /// Every setting unset: sniff the format, use every library default.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The file format. Default [`Format::Auto`], which does not detect CSV - set
    /// [`Format::Csv`] to open one.
    #[must_use]
    pub fn format(mut self, value: Format) -> Self {
        self.format = value;
        self
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
        /// XLS/XLSX/XLSB only. Library default is 512 MiB.
        max_total_decompressed_bytes: i64,
        /// Largest cell value, or buffered CSV record. Library default is 32 MiB.
        max_cell_bytes: i32,
        /// XLS/XLSX/XLSB only. Library default is 128 MiB.
        max_shared_string_bytes: i64,
        /// XLS/XLSX/XLSB only. Library default is 65536.
        max_zip_entries: i32,
        /// XLS/XLSX/XLSB only. Library default is `false`.
        prefetch_decompression: bool,
        /// Deduplicate repeated cell strings. Library default is `false`.
        intern_strings: bool,
        /// XLSX/XLSB only: check each ZIP entry against its CRC-32, so a damaged file fails the read
        /// instead of returning wrong cells. Library default is `true`.
        verify_zip_checksums: bool,
        /// [`open_source`](Self::open_source) only: bytes fetched per `read_at` call. Library
        /// default is 4 MiB; a negative value turns the cache off.
        source_block_size: i32,
        /// [`open_source`](Self::open_source) only: bytes of blocks kept in memory. Library
        /// default is 64 MiB.
        source_cache_bytes: i64,
        /// [`open_reader`](Self::open_reader) with XLSX, XLSB or XLS only: the most bytes the
        /// reader may hold.
        max_buffered_bytes: i64,
    }

    /// Password for an encrypted OOXML workbook. Unset means "not encrypted, or fail with
    /// [`ErrorKind::PasswordRequired`]".
    #[must_use]
    pub fn password(mut self, value: impl Into<String>) -> Self {
        self.password = Some(value.into());
        self
    }

    pub(crate) fn format_raw(&self) -> i32 {
        self.format.raw()
    }

    /// Lowers this into the raw ABI struct, with `struct_size` filled in.
    ///
    /// Returns a guard rather than a bare `XlOpenOptions`: when [`password`](Self::password) is
    /// set, the raw struct's `password` field is a pointer into `self.password`'s bytes, which must
    /// outlive every use of the pointer. The borrow on [`OpenOptionsRaw`] is what makes the
    /// compiler refuse a call site that does not keep `self` alive for as long as the raw struct is
    /// in use.
    ///
    /// `max_cell_bytes`/`intern_strings` lower into both the general and the `csv_*` fields: the
    /// ABI reads the `csv_*` pair for CSV and the other pair for every other format.
    pub(crate) fn to_raw(&self) -> OpenOptionsRaw<'_> {
        let password = self.password.as_deref().map(str::as_bytes);
        let mut raw = XlOpenOptions {
            struct_size: std::mem::size_of::<XlOpenOptions>() as i32,
            csv_sniff_dialect: opt_state(self.csv_sniff_dialect),
            csv_delimiter: opt_number(self.csv_delimiter),
            csv_quote: opt_number(self.csv_quote),
            csv_detect_bom: opt_state(self.csv_detect_bom),
            csv_max_cell_bytes: opt_number(self.max_cell_bytes),
            csv_intern_strings: opt_state(self.intern_strings),
            max_total_decompressed_bytes: opt_number(self.max_total_decompressed_bytes),
            max_cell_bytes: opt_number(self.max_cell_bytes),
            max_shared_string_bytes: opt_number(self.max_shared_string_bytes),
            max_zip_entries: opt_number(self.max_zip_entries),
            prefetch_decompression: opt_state(self.prefetch_decompression),
            intern_strings: opt_state(self.intern_strings),
            password: std::ptr::null(),
            password_len: 0,
            source_cache_bytes: opt_number(self.source_cache_bytes),
            max_buffered_bytes: opt_number(self.max_buffered_bytes),
            source_block_size: opt_number(self.source_block_size),
            verify_zip_checksums: opt_state(self.verify_zip_checksums),
        };
        if let Some(bytes) = password {
            raw.password = bytes.as_ptr();
            raw.password_len = bytes.len() as i32;
        }
        OpenOptionsRaw {
            raw,
            _password: password,
        }
    }

    /// The CSV dialect and cell limit of these options, plus the two per-call settings, as the
    /// aggregation ABI takes them.
    pub(crate) fn to_csv_parallel_raw(
        &self,
        header_row: i32,
        degree_of_parallelism: i32,
    ) -> XlCsvParallelOptions {
        XlCsvParallelOptions {
            struct_size: std::mem::size_of::<XlCsvParallelOptions>() as i32,
            degree_of_parallelism,
            header_row,
            delimiter: opt_number(self.csv_delimiter),
            quote: opt_number(self.csv_quote),
            detect_bom: opt_state(self.csv_detect_bom),
            max_cell_bytes: opt_number(self.max_cell_bytes),
        }
    }
}

/// Holds the raw C struct together with the buffer its `password` pointer refers to. The raw struct
/// must never outlive this guard, and the borrow of `OpenOptions` (carried through `'a`) is what
/// guarantees it.
pub(crate) struct OpenOptionsRaw<'a> {
    raw: XlOpenOptions,
    _password: Option<&'a [u8]>,
}

impl OpenOptionsRaw<'_> {
    /// Pointer to the raw struct, valid for as long as this guard is alive.
    pub(crate) fn as_ptr(&self) -> *const XlOpenOptions {
        &self.raw
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

/// How a sheet is read by [`Sheet::parse`](crate::Sheet::parse),
/// [`Sheet::parse_batches`](crate::Sheet::parse_batches) and their Arrow counterparts.
///
/// ```no_run
/// use excelreader::ReadOptions;
///
/// let options = ReadOptions::new().header_row(2).parallelism(0);
/// ```
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ReadOptions {
    pub(crate) header_row: i32,
    pub(crate) parallelism: i32,
    pub(crate) batch_size: i64,
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self {
            header_row: 1,
            parallelism: 1,
            batch_size: 10_000,
        }
    }
}

impl ReadOptions {
    /// Header in row 1, one thread, 10,000 rows per batch.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// 1-based row whose values name the columns; `0` means no header, and columns are bound by
    /// position.
    #[must_use]
    pub fn header_row(mut self, value: i32) -> Self {
        self.header_row = value;
        self
    }

    /// Threads for a whole-sheet read: `0` every processor, `1` sequential, `n` up to `n`. The
    /// result is the same whatever it is. Batched reads ignore it.
    #[must_use]
    pub fn parallelism(mut self, value: i32) -> Self {
        self.parallelism = value;
        self
    }

    /// Rows per batch for a batched read, `0` meaning one unbounded batch. Whole-sheet reads
    /// ignore it.
    #[must_use]
    pub fn batch_size(mut self, value: i64) -> Self {
        self.batch_size = value;
        self
    }
}

/// Options for [`write_columns`](crate::write_columns), [`write_sheet`](crate::write_sheet) and
/// [`WriterHandle`](crate::WriterHandle). Every setting is unset by default, meaning "use the
/// library default"; set only the ones you want to override.
///
/// ```no_run
/// use excelreader::WriteOptions;
///
/// let options = WriteOptions::new().sheet_name("Dados").use_shared_strings(true);
/// ```
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct WriteOptions {
    format: Format,
    sheet_name: Option<String>,
    csv_delimiter: Option<i32>,
    csv_quote: Option<i32>,
    date1904: Option<bool>,
    use_shared_strings: Option<bool>,
}

impl WriteOptions {
    /// Every setting unset.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// The output format. Default [`Format::Auto`]: inferred from the path's extension
    /// (`.xlsx`/`.xlsb`/`.xls`/`.csv`). Writing to an `impl Write` needs an explicit format.
    #[must_use]
    pub fn format(mut self, value: Format) -> Self {
        self.format = value;
        self
    }

    /// Sheet name. Unset is `"Sheet1"`. Ignored for CSV. Excel's rules (1-31 characters, none of
    /// `: \ / ? * [ ]`) are enforced natively.
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
        /// XLS/XLSB only; ignored for XLSX and CSV. Library default is `false`.
        date1904: bool,
        /// XLSX/XLSB only; ignored for XLS and CSV. Shrinks files with many repeated strings, at
        /// the cost of a string table. Library default is `false`.
        use_shared_strings: bool,
    }

    /// The `XL_FORMAT_*` to write: the explicit format, else one inferred from `path`. With no path
    /// to infer from, `Auto` is rejected here rather than by the native side.
    pub(crate) fn format_for(&self, path: Option<&Path>) -> Result<i32, Error> {
        match (self.format, path) {
            (Format::Auto, Some(path)) => Ok(Format::from_path(path).raw()),
            (Format::Auto, None) => Err(Error::new(
                ErrorKind::InvalidArgument,
                "writing to an io::Write needs an explicit WriteOptions::format",
            )),
            (format, _) => Ok(format.raw()),
        }
    }

    /// Lowers this into the raw ABI struct, with `struct_size` filled in. The struct holds a raw
    /// pointer INTO `self.sheet_name`, so it must not outlive `self`.
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
        let options = OpenOptions::new();
        let guard = options.to_raw();
        let raw = guard.raw;
        assert_eq!(raw.struct_size, std::mem::size_of::<XlOpenOptions>() as i32);
        assert_eq!(raw.csv_sniff_dialect, XL_OPT_DEFAULT);
        assert_eq!(raw.prefetch_decompression, XL_OPT_DEFAULT);
        assert_eq!(raw.csv_delimiter, 0);
        assert_eq!(raw.max_total_decompressed_bytes, 0);
        assert!(
            raw.password.is_null(),
            "no password set means a null pointer"
        );
        assert_eq!(raw.password_len, 0);
        assert_eq!(options.format_raw(), ffi::XL_FORMAT_AUTO);
    }

    #[test]
    fn booleans_lower_to_distinct_true_and_false_states() {
        let options = OpenOptions::new()
            .prefetch_decompression(true)
            .intern_strings(false)
            .verify_zip_checksums(false);
        let guard = options.to_raw();
        let raw = guard.raw;
        assert_eq!(raw.prefetch_decompression, XL_OPT_TRUE);
        assert_eq!(raw.intern_strings, XL_OPT_FALSE);
        assert_eq!(raw.verify_zip_checksums, XL_OPT_FALSE);
        assert_eq!(raw.csv_detect_bom, XL_OPT_DEFAULT);
    }

    #[test]
    fn numbers_pass_through_unchanged() {
        let options = OpenOptions::new()
            .csv_delimiter(b';' as i32)
            .max_zip_entries(1024)
            .max_shared_string_bytes(1 << 20);
        let guard = options.to_raw();
        let raw = guard.raw;
        assert_eq!(raw.csv_delimiter, b';' as i32);
        assert_eq!(raw.max_zip_entries, 1024);
        assert_eq!(raw.max_shared_string_bytes, 1 << 20);
    }

    #[test]
    fn cell_limit_and_interning_reach_both_the_csv_and_the_general_fields() {
        let options = OpenOptions::new().max_cell_bytes(4096).intern_strings(true);
        let guard = options.to_raw();
        let raw = guard.raw;
        assert_eq!(raw.max_cell_bytes, 4096);
        assert_eq!(raw.csv_max_cell_bytes, 4096);
        assert_eq!(raw.intern_strings, XL_OPT_TRUE);
        assert_eq!(raw.csv_intern_strings, XL_OPT_TRUE);
    }

    #[test]
    fn csv_aggregation_options_take_the_dialect_and_cell_limit() {
        let raw = OpenOptions::new()
            .csv_delimiter(b';' as i32)
            .csv_detect_bom(false)
            .max_cell_bytes(1024)
            .to_csv_parallel_raw(1, 4);
        assert_eq!(
            raw.struct_size,
            std::mem::size_of::<XlCsvParallelOptions>() as i32
        );
        assert_eq!(raw.header_row, 1);
        assert_eq!(raw.degree_of_parallelism, 4);
        assert_eq!(raw.delimiter, b';' as i32);
        assert_eq!(raw.quote, 0);
        assert_eq!(raw.detect_bom, XL_OPT_FALSE);
        assert_eq!(raw.max_cell_bytes, 1024);
    }

    #[test]
    fn password_lowers_to_a_pointer_and_length_into_the_owning_options() {
        let options = OpenOptions::new().password("hunter2");
        let guard = options.to_raw();
        let raw = guard.raw;
        assert_eq!(raw.password_len, 7);
        assert!(!raw.password.is_null());
        let bytes = unsafe { std::slice::from_raw_parts(raw.password, raw.password_len as usize) };
        assert_eq!(bytes, b"hunter2");
    }

    /// `to_raw` must not compile down to something that lets the password pointer dangle when the
    /// source string is a temporary rather than a local the caller happens to keep alive.
    #[test]
    fn password_survives_a_temporary_source() {
        let options = OpenOptions::new().password(String::from("hunter2"));
        let guard = options.to_raw();
        let raw = guard.raw;
        let bytes = unsafe { std::slice::from_raw_parts(raw.password, raw.password_len as usize) };
        assert_eq!(bytes, b"hunter2");
    }

    #[test]
    fn write_options_lower_to_the_abi_defaults() {
        let raw = WriteOptions::new().to_raw();
        assert_eq!(
            raw.struct_size,
            std::mem::size_of::<XlWriteOptions>() as i32
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

    #[test]
    fn write_format_is_explicit_or_inferred_from_the_extension() {
        let auto = WriteOptions::new();
        let format = |path: &str| auto.format_for(Some(Path::new(path))).unwrap();
        assert_eq!(format("out.xlsx"), ffi::XL_FORMAT_XLSX);
        assert_eq!(format("out.XLSB"), ffi::XL_FORMAT_XLSB);
        assert_eq!(format("out.xls"), ffi::XL_FORMAT_XLS);
        assert_eq!(format("out.csv"), ffi::XL_FORMAT_CSV);
        assert_eq!(format("out.txt"), ffi::XL_FORMAT_AUTO);
        assert_eq!(format("out"), ffi::XL_FORMAT_AUTO);
        assert_eq!(format("v1.2/report"), ffi::XL_FORMAT_AUTO);

        let error = auto
            .format_for(None)
            .expect_err("a writer has no extension to infer from");
        assert_eq!(error.kind(), ErrorKind::InvalidArgument);

        let csv = WriteOptions::new().format(Format::Csv);
        assert_eq!(
            csv.format_for(Some(Path::new("out.xlsx"))).unwrap(),
            ffi::XL_FORMAT_CSV
        );
        assert_eq!(csv.format_for(None).unwrap(), ffi::XL_FORMAT_CSV);
    }

    /// The C struct's x64 layout is pinned by XL_STATIC_ASSERT in
    /// tests/ExcelReader.NativeSmoke/smoke.c. A repr(C) struct that disagrees would hand the native
    /// side garbage where it expects a length and a pointer.
    #[test]
    fn write_options_match_the_c_struct_size() {
        assert_eq!(std::mem::size_of::<XlWriteOptions>(), 32);
    }

    #[test]
    fn source_layouts_match_the_c_structs() {
        use crate::ffi::{XlSource, XlStream};
        assert_eq!(std::mem::size_of::<XlOpenOptions>(), 112);
        assert_eq!(std::mem::offset_of!(XlOpenOptions, source_cache_bytes), 88);
        assert_eq!(std::mem::offset_of!(XlOpenOptions, max_buffered_bytes), 96);
        assert_eq!(std::mem::offset_of!(XlOpenOptions, source_block_size), 104);
        assert_eq!(
            std::mem::offset_of!(XlOpenOptions, verify_zip_checksums),
            108
        );
        assert_eq!(std::mem::size_of::<XlSource>(), 40);
        assert_eq!(std::mem::size_of::<XlStream>(), 32);
    }
}
