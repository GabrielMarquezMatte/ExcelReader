use std::fmt;

/// A native ExcelReader error: `code` is one of the `XL_*` status constants, `message` is the
/// detail from `xl_last_error_ptr` on the calling thread at the time of failure.
#[derive(Debug, Clone)]
pub struct Error {
    pub code: i32,
    pub message: String,
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "ExcelReader error {}: {}", self.code, self.message)
    }
}

impl std::error::Error for Error {}
