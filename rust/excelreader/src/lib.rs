#![doc = include_str!("../README.md")]
#![doc(
    html_logo_url = "https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png",
    html_favicon_url = "https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/master/assets/icon.png"
)]

mod aggregate;
mod error;
pub(crate) mod ffi;
mod foreach;
mod options;
mod rows;
mod source;
mod temporal;
mod workbook;
mod writer;
mod writer_handle;

#[cfg(feature = "arrow")]
mod arrow;

pub use aggregate::{aggregate_csv, Accumulator, CsvAccumulator, CsvInput};
pub use error::{Error, ErrorKind};
pub use options::{Format, OpenOptions, ReadOptions, WriteOptions};
pub use rows::{AllRows, CellIter, CellRef, CellType, RowCursor, RowRef};
pub use source::Source;
pub use temporal::{Date, Time, Timestamp};
pub use workbook::{
    ColumnType, ExcelMapper, InferredColumn, Sheet, SheetVisibility, TableView, TableViewIter,
    TypedBatches, Workbook,
};
pub use writer::{
    encrypt_package, write_columns, write_columns_to, write_sheet, write_sheet_to, Column,
    ColumnData, ExcelWriter,
};
pub use writer_handle::WriterHandle;

#[cfg(feature = "arrow")]
pub use self::arrow::ArrowBatches;

/// Support for `#[derive(ExcelMapper)]`'s generated code. Not part of the public API: nothing here
/// follows semver.
#[doc(hidden)]
pub mod __private {
    pub use crate::workbook::{
        column_bool, column_date, column_f64, column_i64, column_str, column_time,
        column_timestamp, ColumnBinding, ColumnRef,
    };
    pub use crate::writer::{push_str, set_valid, OwnedColumn, OwnedColumnData};

    #[must_use]
    pub fn invalid_argument(message: String) -> crate::Error {
        crate::Error::new(crate::ErrorKind::InvalidArgument, message)
    }
}
