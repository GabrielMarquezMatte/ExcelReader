//! Row-at-a-time and whole-sheet decoded reads over the C ABI's row APIs.

use crate::{
    Error, XlRowCell, XL_CELL_BOOL, XL_CELL_DATE, XL_CELL_EMPTY, XL_CELL_ERROR, XL_CELL_FORMULA,
    XL_CELL_NUMBER, XL_CELL_STRING, XL_ERROR,
};

/// The kind of a cell, mirroring `XL_CELL_*` in the C ABI.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub enum CellType {
    Empty = XL_CELL_EMPTY,
    String = XL_CELL_STRING,
    Number = XL_CELL_NUMBER,
    Date = XL_CELL_DATE,
    Bool = XL_CELL_BOOL,
    Formula = XL_CELL_FORMULA,
    Error = XL_CELL_ERROR,
}

impl CellType {
    /// Maps a raw `XL_CELL_*` value. `None` for a value this crate does not know, which can only
    /// happen against a newer native library than the one this crate was built against.
    #[must_use]
    pub fn from_raw(value: i32) -> Option<CellType> {
        match value {
            XL_CELL_EMPTY => Some(CellType::Empty),
            XL_CELL_STRING => Some(CellType::String),
            XL_CELL_NUMBER => Some(CellType::Number),
            XL_CELL_DATE => Some(CellType::Date),
            XL_CELL_BOOL => Some(CellType::Bool),
            XL_CELL_FORMULA => Some(CellType::Formula),
            XL_CELL_ERROR => Some(CellType::Error),
            _ => None,
        }
    }
}

/// One cell, borrowing its bytes from the row that produced it.
#[derive(Debug, Clone, Copy)]
pub struct CellRef<'a> {
    pub column: i32,
    pub cell_type: CellType,
    value: &'a [u8],
}

impl<'a> CellRef<'a> {
    /// The raw UTF-8 bytes as stored. A `Date` cell carries an Excel serial number as text.
    #[must_use]
    pub fn as_bytes(&self) -> &'a [u8] {
        self.value
    }

    /// The value as a string. Fails when the bytes are not valid UTF-8, which a well-formed
    /// workbook never produces.
    pub fn as_str(&self) -> Result<&'a str, Error> {
        std::str::from_utf8(self.value)
            .map_err(|err| Error::from_status(XL_ERROR, format!("cell value is not valid UTF-8: {err}")))
    }
}

/// Where a `RowRef`'s cells live. Blob rows come from `xl_next_row`, decoded rows from
/// `xl_read_all_decoded`.
#[allow(dead_code)]
#[derive(Debug, Clone, Copy)]
enum RowBacking<'a> {
    /// The bytes AFTER the leading `int32 cell_count`.
    Blob(&'a [u8]),
    Decoded(&'a [XlRowCell]),
}

/// One row, borrowing from whichever buffer produced it.
#[derive(Debug, Clone, Copy)]
pub struct RowRef<'a> {
    backing: RowBacking<'a>,
    len: usize,
}

impl<'a> RowRef<'a> {
    /// Parses the leading cell count off a `xl_next_row` blob. `None` when the blob is too short to
    /// hold even that count.
    #[allow(dead_code)]
    pub(crate) fn from_blob(blob: &'a [u8]) -> Option<RowRef<'a>> {
        let count = read_i32(blob, 0)?;
        if count < 0 {
            return None;
        }
        Some(RowRef {
            backing: RowBacking::Blob(&blob[4..]),
            len: count as usize,
        })
    }

    /// Wraps the cells of one `XlRow`.
    ///
    /// # Safety
    /// `cells` must point to `count` initialized `XlRowCell` values whose `value` pointers stay
    /// valid for `'a` — that is, until `xl_free_rows` releases the enclosing `XlRows`.
    #[allow(dead_code)]
    pub(crate) unsafe fn from_decoded(cells: *const XlRowCell, count: i32) -> RowRef<'a> {
        let len = if count > 0 { count as usize } else { 0 };
        let slice = if len == 0 || cells.is_null() {
            &[][..]
        } else {
            unsafe { std::slice::from_raw_parts(cells, len) }
        };
        RowRef { backing: RowBacking::Decoded(slice), len }
    }

    #[must_use]
    pub fn len(&self) -> usize {
        self.len
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.len == 0
    }

    /// The cell at `index`. For a row from `RowCursor` this walks the blob from the start, so it is
    /// O(index); prefer `iter()` when reading a whole row. A row from `DecodedRows` indexes
    /// directly.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<CellRef<'a>> {
        if index >= self.len {
            return None;
        }
        match self.backing {
            RowBacking::Decoded(cells) => cell_from_decoded(&cells[index]),
            RowBacking::Blob(_) => self.iter().nth(index),
        }
    }

    #[must_use]
    pub fn iter(&self) -> CellIter<'a> {
        CellIter { backing: self.backing, len: self.len, index: 0, offset: 0 }
    }
}

impl<'a> IntoIterator for RowRef<'a> {
    type Item = CellRef<'a>;
    type IntoIter = CellIter<'a>;

    fn into_iter(self) -> CellIter<'a> {
        self.iter()
    }
}

/// Walks a row's cells left to right.
#[derive(Debug, Clone)]
pub struct CellIter<'a> {
    backing: RowBacking<'a>,
    len: usize,
    index: usize,
    offset: usize,
}

impl<'a> Iterator for CellIter<'a> {
    type Item = CellRef<'a>;

    fn next(&mut self) -> Option<CellRef<'a>> {
        if self.index >= self.len {
            return None;
        }
        match self.backing {
            RowBacking::Decoded(cells) => {
                let cell = cell_from_decoded(&cells[self.index])?;
                self.index += 1;
                Some(cell)
            }
            RowBacking::Blob(blob) => {
                let (cell, next) = cell_from_blob(blob, self.offset)?;
                self.offset = next;
                self.index += 1;
                Some(cell)
            }
        }
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.len - self.index;
        (0, Some(remaining))
    }
}

fn read_i32(bytes: &[u8], offset: usize) -> Option<i32> {
    let end = offset.checked_add(4)?;
    let slice = bytes.get(offset..end)?;
    Some(i32::from_le_bytes(slice.try_into().ok()?))
}

/// Decodes the cell starting at `offset`, returning it with the offset of the next one. `None` for
/// a truncated or malformed blob, which is why every read here is bounds-checked rather than
/// trusting the declared cell count.
fn cell_from_blob(blob: &[u8], offset: usize) -> Option<(CellRef<'_>, usize)> {
    let column = read_i32(blob, offset)?;
    let raw_type = read_i32(blob, offset + 4)?;
    let value_len = read_i32(blob, offset + 8)?;
    if value_len < 0 {
        return None;
    }
    let start = offset.checked_add(12)?;
    let end = start.checked_add(value_len as usize)?;
    let value = blob.get(start..end)?;
    let cell_type = CellType::from_raw(raw_type)?;
    Some((CellRef { column, cell_type, value }, end))
}

fn cell_from_decoded(raw: &XlRowCell) -> Option<CellRef<'_>> {
    let cell_type = CellType::from_raw(raw.cell_type)?;
    // from_raw_parts requires a non-null, aligned pointer even for a zero-length slice.
    let value = if raw.value.is_null() || raw.value_len <= 0 {
        &[][..]
    } else {
        unsafe { std::slice::from_raw_parts(raw.value, raw.value_len as usize) }
    };
    Some(CellRef { column: raw.column, cell_type, value })
}

/// Rows are usually well under this; it only sets how often the first oversized row costs a retry.
const INITIAL_ROW_BUFFER: usize = 64 * 1024;

/// A row-at-a-time reader over a workbook's current sheet, holding one reusable buffer.
///
/// Not an `Iterator`: each row borrows the buffer that the next call overwrites, which
/// `Iterator::next` cannot express. One row is alive at a time, enforced by the borrow checker.
pub struct RowCursor<'w> {
    handle: *mut crate::XlWorkbook,
    buffer: Vec<u8>,
    written: usize,
    workbook: std::marker::PhantomData<&'w mut crate::workbook::Workbook>,
}

impl<'w> RowCursor<'w> {
    pub(crate) fn new(handle: *mut crate::XlWorkbook) -> RowCursor<'w> {
        RowCursor {
            handle,
            buffer: vec![0; INITIAL_ROW_BUFFER],
            written: 0,
            workbook: std::marker::PhantomData,
        }
    }

    /// Advances to the next row. `None` at end of sheet.
    ///
    /// On `XL_BUFFER_TOO_SMALL` the native side holds the row until it fits, so growing the buffer
    /// and retrying loses nothing.
    pub fn next_row(&mut self) -> Option<Result<RowRef<'_>, Error>> {
        loop {
            let mut written: i32 = 0;
            let capacity = i32::try_from(self.buffer.len()).unwrap_or(i32::MAX);
            let status = unsafe {
                crate::xl_next_row(self.handle, self.buffer.as_mut_ptr(), capacity, &mut written)
            };

            match status {
                crate::XL_OK => {
                    self.written = if written > 0 { written as usize } else { 0 };
                    let blob = &self.buffer[..self.written];
                    return Some(
                        RowRef::from_blob(blob).ok_or_else(|| {
                            Error::from_status(XL_ERROR, "native returned a malformed row blob".to_string())
                        }),
                    );
                }
                crate::XL_EOF => return None,
                crate::XL_BUFFER_TOO_SMALL => {
                    let needed = if written > 0 { written as usize } else { self.buffer.len() * 2 };
                    if needed <= self.buffer.len() {
                        return Some(Err(Error::from_status(
                            XL_ERROR,
                            "native asked for a buffer no larger than the current one".to_string(),
                        )));
                    }
                    self.buffer.resize(needed, 0);
                }
                other => return Some(Err(crate::workbook::last_error(other))),
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Builds a row blob in the wire format documented on `xl_next_row`:
    /// `int32 cell_count`, then per cell `int32 column, int32 type, int32 value_len, bytes`.
    fn blob(cells: &[(i32, i32, &str)]) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(cells.len() as i32).to_le_bytes());
        for (column, cell_type, value) in cells {
            out.extend_from_slice(&column.to_le_bytes());
            out.extend_from_slice(&cell_type.to_le_bytes());
            out.extend_from_slice(&(value.len() as i32).to_le_bytes());
            out.extend_from_slice(value.as_bytes());
        }
        out
    }

    #[test]
    fn decodes_a_blob_row() {
        let bytes = blob(&[(0, XL_CELL_STRING, "hello"), (2, XL_CELL_NUMBER, "42")]);
        let row = RowRef::from_blob(&bytes).expect("well-formed blob");

        assert_eq!(row.len(), 2);
        assert!(!row.is_empty());

        let first = row.get(0).expect("cell 0");
        assert_eq!(first.column, 0);
        assert_eq!(first.cell_type, CellType::String);
        assert_eq!(first.as_str().unwrap(), "hello");

        let second = row.get(1).expect("cell 1");
        assert_eq!(second.column, 2);
        assert_eq!(second.cell_type, CellType::Number);
        assert_eq!(second.as_str().unwrap(), "42");

        assert!(row.get(2).is_none());
    }

    #[test]
    fn iterates_in_order() {
        let bytes = blob(&[(0, XL_CELL_STRING, "a"), (1, XL_CELL_STRING, "b"), (2, XL_CELL_STRING, "c")]);
        let row = RowRef::from_blob(&bytes).expect("well-formed blob");
        let values: Vec<&str> = row.iter().map(|cell| cell.as_str().unwrap()).collect();
        assert_eq!(values, ["a", "b", "c"]);
    }

    #[test]
    fn empty_row_decodes() {
        let bytes = blob(&[]);
        let row = RowRef::from_blob(&bytes).expect("well-formed blob");
        assert_eq!(row.len(), 0);
        assert!(row.is_empty());
        assert_eq!(row.iter().count(), 0);
    }

    #[test]
    fn truncated_blob_is_rejected_not_panicked() {
        // One cell declared, but the value bytes are cut short.
        let mut bytes = blob(&[(0, XL_CELL_STRING, "hello")]);
        bytes.truncate(bytes.len() - 3);
        let row = RowRef::from_blob(&bytes).expect("header is intact");
        assert!(row.get(0).is_none());
    }

    #[test]
    fn unknown_cell_type_is_none() {
        assert_eq!(CellType::from_raw(99), None);
        assert_eq!(CellType::from_raw(XL_CELL_ERROR), Some(CellType::Error));
    }
}
