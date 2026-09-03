//! Writing a columnar table to a workbook file, through the ABI's single `xl_write_typed` export.
//!
//! Two layers sit here. [`write_columns`] takes buffers the caller already owns and hands them
//! straight to the ABI - the borrow is the whole point, and the lifetimes on [`Column`] are what
//! turn the ABI's prose contract ("every buffer is borrowed for the duration of the call") into
//! something the compiler enforces. [`write_sheet`] sits on top for callers holding a slice of
//! structs, and pays exactly one transpose to get them into columns.

use crate::workbook::{check, check_abi_version};
use crate::{
    Error, WriteOptions, XlColumn, XlColumnSpec, XlTable, XL_FORMAT_AUTO, XL_FORMAT_CSV,
    XL_FORMAT_XLS, XL_FORMAT_XLSB, XL_FORMAT_XLSX, XL_INVALID_ARGUMENT, XL_T_BOOL, XL_T_DATE,
    XL_T_F64, XL_T_I64, XL_T_STRING, XL_T_TIME, XL_T_TIMESTAMP,
};
use std::os::raw::c_void;

/// One column's buffers, borrowed from the caller. Each variant's slice is that column type's
/// exact wire layout (see `excelreader.h`), so nothing is converted on the way out.
pub enum ColumnData<'a> {
    /// `offsets` has `rows + 1` entries into `data`, which is every row's UTF-8 bytes
    /// concatenated. Unlike a table returned by the reader, `data` need not be interior to
    /// `offsets` here.
    Str {
        offsets: &'a [i32],
        data: &'a [u8],
    },
    I64(&'a [i64]),
    F64(&'a [f64]),
    /// One byte per row, 0 or 1 - NOT a bit-packed bitmap.
    Bool(&'a [u8]),
    /// Days since 1970-01-01.
    Date(&'a [i32]),
    /// Microseconds since midnight.
    Time(&'a [i64]),
    /// Microseconds since 1970-01-01T00:00:00Z.
    Timestamp(&'a [i64]),
}

impl ColumnData<'_> {
    #[must_use]
    pub fn len(&self) -> i64 {
        match self {
            // An offsets array of n + 1 entries describes n rows; an empty one describes none.
            ColumnData::Str { offsets, .. } => (offsets.len().max(1) - 1) as i64,
            ColumnData::I64(values) | ColumnData::Time(values) | ColumnData::Timestamp(values) => {
                values.len() as i64
            }
            ColumnData::F64(values) => values.len() as i64,
            ColumnData::Bool(values) => values.len() as i64,
            ColumnData::Date(values) => values.len() as i64,
        }
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }

    /// The `XL_T_*` tag this variant writes as.
    #[must_use]
    pub fn xl_type(&self) -> i32 {
        match self {
            ColumnData::Str { .. } => XL_T_STRING,
            ColumnData::I64(_) => XL_T_I64,
            ColumnData::F64(_) => XL_T_F64,
            ColumnData::Bool(_) => XL_T_BOOL,
            ColumnData::Date(_) => XL_T_DATE,
            ColumnData::Time(_) => XL_T_TIME,
            ColumnData::Timestamp(_) => XL_T_TIMESTAMP,
        }
    }

    fn pointers(&self) -> (*const c_void, *const u8, i64) {
        match self {
            ColumnData::Str { offsets, data } => (
                offsets.as_ptr().cast::<c_void>(),
                data.as_ptr(),
                data.len() as i64,
            ),
            ColumnData::I64(v) | ColumnData::Time(v) | ColumnData::Timestamp(v) => {
                (v.as_ptr().cast::<c_void>(), std::ptr::null(), 0)
            }
            ColumnData::F64(v) => (v.as_ptr().cast::<c_void>(), std::ptr::null(), 0),
            ColumnData::Bool(v) => (v.as_ptr().cast::<c_void>(), std::ptr::null(), 0),
            ColumnData::Date(v) => (v.as_ptr().cast::<c_void>(), std::ptr::null(), 0),
        }
    }
}

/// One input column: a name, its buffers, and an optional validity bitmap.
pub struct Column<'a> {
    /// The header text. `None` in EVERY column means no header row is written; mixing `Some` and
    /// `None` across the set is an error, not a partial header.
    pub name: Option<&'a str>,
    pub data: ColumnData<'a>,
    /// LSB-first bitmap, bit `r` set = row `r` is valid. `None` = the column has no nulls.
    pub validity: Option<&'a [u8]>,
}

impl Column<'_> {
    fn to_raw(&self) -> XlColumn {
        let (values, data, data_len) = self.data.pointers();
        XlColumn {
            r#type: self.data.xl_type(),
            length: self.data.len(),
            values,
            validity: self.validity.map_or(std::ptr::null(), <[u8]>::as_ptr),
            data,
            data_len,
        }
    }
}

/// Infers an `XL_FORMAT_*` from a path's extension, case-insensitively. Returns
/// [`XL_FORMAT_AUTO`] when the extension is absent or unrecognized - and since `xl_write_typed`
/// rejects `AUTO` with a message of its own, an unrecognized path fails the write rather than
/// silently picking a format.
#[must_use]
pub fn format_from_path(path: &str) -> i32 {
    let name = path.rsplit(['/', '\\']).next().unwrap_or(path);
    let Some((_, extension)) = name.rsplit_once('.') else {
        return XL_FORMAT_AUTO;
    };
    match extension.to_ascii_lowercase().as_str() {
        "xlsx" => XL_FORMAT_XLSX,
        "xlsb" => XL_FORMAT_XLSB,
        "xls" => XL_FORMAT_XLS,
        "csv" => XL_FORMAT_CSV,
        _ => XL_FORMAT_AUTO,
    }
}

fn invalid(message: String) -> Error {
    Error::from_status(XL_INVALID_ARGUMENT, message)
}

/// Returns the row count every column agreed on, or the first problem found.
///
/// Runs to completion before anything reaches the native side. The bitmap length check is the one
/// that must live here and cannot be delegated: `xl_write_typed` takes `validity` without a length
/// and reads `(rows + 7) / 8` bytes on trust, so a short slice is a buffer overrun the ABI has no
/// way to catch.
fn validate(columns: &[Column<'_>]) -> Result<i64, Error> {
    let Some(first) = columns.first() else {
        return Err(invalid(
            "write_columns needs at least one column.".to_string(),
        ));
    };
    let rows = first.data.len();
    let has_header = first.name.is_some();

    for (index, column) in columns.iter().enumerate() {
        if column.data.len() != rows {
            return Err(invalid(format!(
                "every column must have the same length; column 0 has {rows} rows but column \
                 {index} has {}",
                column.data.len()
            )));
        }
        if column.name.is_some() != has_header {
            return Err(invalid(format!(
                "every column must have a name, or none may - xl_write_typed cannot write a \
                 partial header row (column {index})"
            )));
        }
        if let Some(bitmap) = column.validity {
            let needed = (rows as usize).div_ceil(8);
            if bitmap.len() < needed {
                return Err(invalid(format!(
                    "the validity bitmap is {} bytes, but {rows} rows need {needed} (column \
                     {index})",
                    bitmap.len()
                )));
            }
        }
        if let ColumnData::Str { offsets, data } = &column.data {
            if data.len() > i32::MAX as usize {
                return Err(invalid(format!(
                    "the string blob is larger than 2 GiB, which int32 offsets cannot address \
                     (column {index})"
                )));
            }
            if offsets.len() as i64 != rows + 1 {
                return Err(invalid(format!(
                    "a string column needs {} offsets for {rows} rows; column {index} has {}",
                    rows + 1,
                    offsets.len()
                )));
            }
        }
    }
    Ok(rows)
}

/// Writes `columns` to `path` as a single sheet, then closes the file. One-shot: no writer handle
/// exists before or after, and every buffer reachable from `columns` and `options` is borrowed for
/// the duration of the call and never freed by the native library.
///
/// `format` must be one of `XL_FORMAT_XLS`/`XLSX`/`XLSB`/`CSV`. [`XL_FORMAT_AUTO`] is an error: a
/// file being created has no signature bytes to sniff. Use [`format_from_path`] to infer one.
///
/// On failure the destination may exist and be incomplete; cleaning it up is the caller's.
pub fn write_columns(
    path: &str,
    format: i32,
    columns: &[Column<'_>],
    options: Option<&WriteOptions>,
) -> Result<(), Error> {
    check_abi_version()?;
    let row_count = validate(columns)?;

    // Three parallel arrays that must outlive the call: each spec's `names` points at one slot of
    // `names`, and its `name_lens` at one slot of `name_lens`. A temporary would dangle.
    let names: Vec<*const u8> = columns
        .iter()
        .map(|c| c.name.map_or(std::ptr::null(), str::as_ptr))
        .collect();
    let name_lens: Vec<i32> = columns
        .iter()
        .map(|c| c.name.map_or(0, |n| n.len() as i32))
        .collect();
    let specs: Vec<XlColumnSpec> = columns
        .iter()
        .enumerate()
        .map(|(index, column)| XlColumnSpec {
            names: &names[index],
            name_lens: &name_lens[index],
            // Exactly one name per write spec, or none: the ABI rejects a spec carrying an alias
            // list, which only exists to resolve a header on the way IN.
            name_count: i32::from(column.name.is_some()),
            index: 0,
            r#type: column.data.xl_type(),
            nullable: 0,
        })
        .collect();
    let raw_columns: Vec<XlColumn> = columns.iter().map(Column::to_raw).collect();

    let table = XlTable {
        column_count: columns.len() as i32,
        row_count,
        // `columns` is `*mut` in the C struct only because the reader fills one in; the writer
        // takes the table as `const` and never writes through it.
        columns: raw_columns.as_ptr().cast_mut(),
    };
    // Lowered here rather than by the caller: the raw struct holds a pointer into
    // `options.sheet_name`, and this borrow provably covers the FFI call below.
    let raw_options = options.map(WriteOptions::to_raw);
    let options_ptr = crate::options::ptr_or_null(&raw_options);

    let status = unsafe {
        crate::xl_write_typed(
            path.as_ptr(),
            path.len() as i32,
            format,
            specs.as_ptr(),
            &table,
            options_ptr,
        )
    };
    check(status)
}

/// In-memory equivalent of [`write_columns`]: same validation and column lowering, but the
/// workbook is built in memory and returned as bytes instead of being written to a path - so,
/// unlike [`write_columns`], there is no path to infer a format from and `format` is always
/// required.
///
/// # Errors
/// Anything [`write_columns`] reports.
pub fn write_columns_to_memory(
    format: i32,
    columns: &[Column<'_>],
    options: Option<&WriteOptions>,
) -> Result<Vec<u8>, Error> {
    check_abi_version()?;
    let row_count = validate(columns)?;

    let names: Vec<*const u8> = columns
        .iter()
        .map(|c| c.name.map_or(std::ptr::null(), str::as_ptr))
        .collect();
    let name_lens: Vec<i32> = columns
        .iter()
        .map(|c| c.name.map_or(0, |n| n.len() as i32))
        .collect();
    let specs: Vec<XlColumnSpec> = columns
        .iter()
        .enumerate()
        .map(|(index, column)| XlColumnSpec {
            names: &names[index],
            name_lens: &name_lens[index],
            name_count: i32::from(column.name.is_some()),
            index: 0,
            r#type: column.data.xl_type(),
            nullable: 0,
        })
        .collect();
    let raw_columns: Vec<XlColumn> = columns.iter().map(Column::to_raw).collect();

    let table = XlTable {
        column_count: columns.len() as i32,
        row_count,
        columns: raw_columns.as_ptr().cast_mut(),
    };
    let raw_options = options.map(WriteOptions::to_raw);
    let options_ptr = crate::options::ptr_or_null(&raw_options);

    let mut buffer = crate::XlBuffer {
        data: std::ptr::null_mut(),
        len: 0,
    };
    let status = unsafe {
        crate::xl_write_typed_to_memory(format, specs.as_ptr(), &table, options_ptr, &mut buffer)
    };
    check(status)?;
    Ok(crate::workbook::buffer_to_vec(buffer))
}

/// Wraps a finished plaintext XLSX/XLSB package at `package_path` in an agile-encrypted (ECMA-376
/// 4.4) CFB container, written to `destination_path` (overwriting an existing file). The result
/// opens with the same password via [`crate::OpenOptions::password`]. Encryption parameters are
/// fixed at Excel's own defaults - there are no options.
///
/// `package_path` is read twice, so it must already be a finished file (write it with
/// [`write_columns`]/[`write_sheet`] first).
///
/// # Errors
/// `package_path` does not exist, is not readable, or is not a valid plaintext OOXML package;
/// `password` is empty; or the usual file I/O failures on either path.
pub fn encrypt_package(package_path: &str, destination_path: &str, password: &str) -> Result<(), Error> {
    check_abi_version()?;
    let status = unsafe {
        crate::xl_encrypt_package(
            package_path.as_ptr(),
            package_path.len() as i32,
            destination_path.as_ptr(),
            destination_path.len() as i32,
            password.as_ptr(),
            password.len() as i32,
        )
    };
    check(status)
}

/// The owning twin of [`ColumnData`], produced by [`ExcelWriter::to_columns`]. A transposed range
/// of structs has to own its columns somewhere; this is that somewhere.
pub enum OwnedColumnData {
    Str { offsets: Vec<i32>, data: Vec<u8> },
    I64(Vec<i64>),
    F64(Vec<f64>),
    Bool(Vec<u8>),
    Date(Vec<i32>),
    Time(Vec<i64>),
    Timestamp(Vec<i64>),
}

/// One owned column. `name` is `&'static str` rather than `String` because it always comes from a
/// literal in a `#[excel(name = "...")]` attribute - keeping it borrowed means transposing a
/// million rows allocates nothing for names.
pub struct OwnedColumn {
    pub name: Option<&'static str>,
    pub data: OwnedColumnData,
    /// LSB-first bitmap. `None` = the column has no nulls.
    pub validity: Option<Vec<u8>>,
}

impl OwnedColumn {
    /// Borrows this column in the shape [`write_columns`] takes.
    #[must_use]
    pub fn as_column(&self) -> Column<'_> {
        let data = match &self.data {
            OwnedColumnData::Str { offsets, data } => ColumnData::Str { offsets, data },
            OwnedColumnData::I64(v) => ColumnData::I64(v),
            OwnedColumnData::F64(v) => ColumnData::F64(v),
            OwnedColumnData::Bool(v) => ColumnData::Bool(v),
            OwnedColumnData::Date(v) => ColumnData::Date(v),
            OwnedColumnData::Time(v) => ColumnData::Time(v),
            OwnedColumnData::Timestamp(v) => ColumnData::Timestamp(v),
        };
        Column {
            name: self.name,
            data,
            validity: self.validity.as_deref(),
        }
    }
}

/// Whether appending `added` bytes to a blob already `current` bytes long would push the next
/// offset past what an `int32` can hold. Split out of `push_str` so the arithmetic is testable
/// without materializing a 2 GiB buffer.
fn offset_ceiling_exceeded(current: usize, added: usize) -> bool {
    current.saturating_add(added) > i32::MAX as usize
}

/// Appends one string to a column's offsets/data pair.
///
/// This is where the `int32` offset overflow is caught, and it has to be caught HERE rather than
/// after the fact: once `data` has grown past `i32::MAX` the offset that would record it has
/// already wrapped, and a wrapped offset is indistinguishable from a real one.
///
/// # Errors
/// When appending `value` would push the blob past `i32::MAX` bytes.
pub fn push_str(offsets: &mut Vec<i32>, data: &mut Vec<u8>, value: &str) -> Result<(), Error> {
    if offset_ceiling_exceeded(data.len(), value.len()) {
        return Err(invalid(
            "a string column exceeds 2 GiB, which int32 offsets cannot address.".to_string(),
        ));
    }
    data.extend_from_slice(value.as_bytes());
    offsets.push(data.len() as i32);
    Ok(())
}

/// Marks row `row` valid in an LSB-first bitmap.
///
/// # Panics
/// If `validity` is shorter than `row / 8 + 1` bytes - a caller-side sizing bug, not recoverable
/// input.
pub fn set_valid(validity: &mut [u8], row: usize) {
    validity[row / 8] |= 1 << (row % 8);
}

/// Implemented by any struct [`write_sheet`] can write. Derive it with
/// `#[derive(ExcelMapper)]`, which emits this alongside the reading half, or write it by hand.
pub trait ExcelWriter: Sized {
    /// Transposes `rows` into one column per field, in field order.
    ///
    /// # Errors
    /// Only for values the ABI cannot represent: a string column past 2 GiB, or an integer that
    /// does not fit `i64`. Ordinary data never fails here.
    fn to_columns(rows: &[Self]) -> Result<Vec<OwnedColumn>, Error>;
}

/// Writes `rows` to `path` as a single sheet, using the same field mapping
/// [`parse_sheet`](crate::workbook::parse_sheet) reads with.
///
/// `rows` is walked ONCE and each field appended to its own column buffer. That transpose is the
/// only copy this makes - it is what the ABI's columnar shape costs a row-shaped caller. If you
/// already hold columnar buffers, call [`write_columns`] and pay nothing.
///
/// # Errors
/// Anything [`ExcelWriter::to_columns`] or [`write_columns`] reports.
pub fn write_sheet<T: ExcelWriter>(
    path: &str,
    format: i32,
    rows: &[T],
    options: Option<&WriteOptions>,
) -> Result<(), Error> {
    let owned = T::to_columns(rows)?;
    let borrowed: Vec<Column<'_>> = owned.iter().map(OwnedColumn::as_column).collect();
    write_columns(path, format, &borrowed, options)
}

/// In-memory equivalent of [`write_sheet`]: same transpose, but returns bytes instead of writing
/// to a path - see [`write_columns_to_memory`] for why `format` is always required here.
///
/// # Errors
/// Anything [`ExcelWriter::to_columns`] or [`write_columns_to_memory`] reports.
pub fn write_sheet_to_memory<T: ExcelWriter>(
    format: i32,
    rows: &[T],
    options: Option<&WriteOptions>,
) -> Result<Vec<u8>, Error> {
    let owned = T::to_columns(rows)?;
    let borrowed: Vec<Column<'_>> = owned.iter().map(OwnedColumn::as_column).collect();
    write_columns_to_memory(format, &borrowed, options)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_is_inferred_from_the_extension() {
        assert_eq!(format_from_path("out.xlsx"), XL_FORMAT_XLSX);
        assert_eq!(format_from_path("out.XLSB"), XL_FORMAT_XLSB);
        assert_eq!(format_from_path("out.xls"), XL_FORMAT_XLS);
        assert_eq!(format_from_path("out.csv"), XL_FORMAT_CSV);
        assert_eq!(format_from_path("out.txt"), XL_FORMAT_AUTO);
        assert_eq!(format_from_path("out"), XL_FORMAT_AUTO);
        // A dot in a directory name is not an extension.
        assert_eq!(format_from_path("v1.2/report"), XL_FORMAT_AUTO);
    }

    #[test]
    fn validate_rejects_a_short_validity_bitmap() {
        let values = [1i64, 2, 3, 4, 5, 6, 7, 8, 9];
        let bitmap = [0u8]; // 9 rows need 2 bytes
        let columns = [Column {
            name: Some("a"),
            data: ColumnData::I64(&values),
            validity: Some(&bitmap),
        }];
        let error = validate(&columns).expect_err("a 1-byte bitmap cannot cover 9 rows");
        assert_eq!(error.code(), XL_INVALID_ARGUMENT);
    }

    #[test]
    fn validate_accepts_an_exactly_sized_validity_bitmap() {
        let values = [1i64; 8];
        let bitmap = [0xFFu8];
        let columns = [Column {
            name: Some("a"),
            data: ColumnData::I64(&values),
            validity: Some(&bitmap),
        }];
        assert_eq!(validate(&columns).expect("8 rows fit in 1 byte"), 8);
    }

    #[test]
    fn validate_rejects_a_string_column_with_the_wrong_offset_count() {
        let offsets = [0i32, 3];
        let columns = [Column {
            name: Some("a"),
            data: ColumnData::Str {
                offsets: &offsets,
                data: b"abc",
            },
            validity: None,
        }];
        // One row, so the offsets array is right - but claiming two rows' worth is not.
        assert_eq!(validate(&columns).expect("1 row, 2 offsets"), 1);

        let bad = [0i32, 3, 6, 9];
        let mixed = [
            Column {
                name: Some("a"),
                data: ColumnData::I64(&[1, 2]),
                validity: None,
            },
            Column {
                name: Some("b"),
                data: ColumnData::Str {
                    offsets: &bad,
                    data: b"abcdefghi",
                },
                validity: None,
            },
        ];
        assert!(
            validate(&mixed).is_err(),
            "3 rows next to a 2-row column must be rejected"
        );
    }
    #[test]
    fn push_str_appends_bytes_and_one_offset_per_value() {
        let mut offsets = vec![0i32];
        let mut data = Vec::new();
        push_str(&mut offsets, &mut data, "uma").expect("a short string must fit");
        push_str(&mut offsets, &mut data, "").expect("an empty string must fit");
        push_str(&mut offsets, &mut data, "duas").expect("a short string must fit");
        assert_eq!(offsets, vec![0, 3, 3, 7]);
        assert_eq!(data, b"umaduas");
    }

    /// The ceiling itself, tested without allocating 2 GiB. `push_str` is a two-liner around this
    /// predicate; the predicate is where the arithmetic that could be wrong lives.
    #[test]
    fn the_offset_ceiling_is_exceeded_exactly_at_int32_max() {
        let ceiling = i32::MAX as usize;
        assert!(!offset_ceiling_exceeded(0, 0));
        assert!(
            !offset_ceiling_exceeded(ceiling - 1, 1),
            "landing exactly on i32::MAX fits"
        );
        assert!(
            offset_ceiling_exceeded(ceiling, 1),
            "one byte past i32::MAX does not"
        );
        assert!(offset_ceiling_exceeded(ceiling - 1, 2));
    }

    #[test]
    fn set_valid_sets_the_lsb_first_bit_for_a_row() {
        let mut bitmap = vec![0u8; 2];
        set_valid(&mut bitmap, 0);
        set_valid(&mut bitmap, 2);
        set_valid(&mut bitmap, 9);
        assert_eq!(bitmap, vec![0b0000_0101, 0b0000_0010]);
    }
}
