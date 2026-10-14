use std::fmt;

use crate::ffi;

/// What went wrong, for callers that branch on the failure. The human-readable detail is
/// [`Error::message`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ErrorKind {
    /// An argument was rejected before any work was done: a bad option value, an out-of-range
    /// index, mismatched column lengths, a path that is not UTF-8.
    InvalidArgument,
    /// The workbook is encrypted and no password was supplied.
    PasswordRequired,
    /// The supplied password did not match the workbook's verifier.
    PasswordIncorrect,
    /// A [`CsvAccumulator`](crate::CsvAccumulator) callback returned this code, passed through
    /// verbatim.
    Callback(i32),
    /// Anything else: an I/O failure, a malformed or unsupported file, a resource limit.
    Other,
}

/// An ExcelReader error: an [`ErrorKind`] plus the detail message the native library (or this
/// crate) reported at the time of failure.
#[derive(Clone, Debug)]
pub struct Error {
    kind: ErrorKind,
    message: String,
}

impl Error {
    pub(crate) fn new(kind: ErrorKind, message: impl Into<String>) -> Self {
        Error {
            kind,
            message: message.into(),
        }
    }

    /// The one place a native status code becomes an [`ErrorKind`]. Positive codes only ever come
    /// from an aggregation callback; every library status is `<= 0`.
    pub(crate) fn from_status(code: i32, message: String) -> Self {
        let kind = match code {
            ffi::XL_INVALID_ARGUMENT => ErrorKind::InvalidArgument,
            ffi::XL_STATUS_PASSWORD_REQUIRED => ErrorKind::PasswordRequired,
            ffi::XL_STATUS_PASSWORD_INCORRECT => ErrorKind::PasswordIncorrect,
            code if code > 0 => ErrorKind::Callback(code),
            _ => ErrorKind::Other,
        };
        Error { kind, message }
    }

    pub(crate) fn io(error: &std::io::Error) -> Self {
        Error::new(ErrorKind::Other, error.to_string())
    }

    #[must_use]
    pub fn kind(&self) -> ErrorKind {
        self.kind
    }

    /// The detail reported at the time of failure.
    #[must_use]
    pub fn message(&self) -> &str {
        &self.message
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "ExcelReader error ({:?}): {}", self.kind, self.message)
    }
}

impl std::error::Error for Error {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn statuses_map_onto_kinds() {
        let kind = |code| Error::from_status(code, String::new()).kind();
        assert_eq!(kind(ffi::XL_INVALID_ARGUMENT), ErrorKind::InvalidArgument);
        assert_eq!(
            kind(ffi::XL_STATUS_PASSWORD_REQUIRED),
            ErrorKind::PasswordRequired
        );
        assert_eq!(
            kind(ffi::XL_STATUS_PASSWORD_INCORRECT),
            ErrorKind::PasswordIncorrect
        );
        assert_eq!(kind(42), ErrorKind::Callback(42));
        assert_eq!(kind(-5), ErrorKind::Other);
    }
}
