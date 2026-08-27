use std::fmt;

/// A native ExcelReader error. `code` is one of the `XL_*` status constants; `message` is the
/// detail from `xl_last_error_ptr` on the calling thread at the time of failure.
///
/// [`PasswordRequired`](Error::PasswordRequired) and [`PasswordIncorrect`](Error::PasswordIncorrect)
/// get their own variants - rather than staying folded into a generic `code`/`message` pair - so
/// callers can `match`/`matches!` on "needs a password" without hardcoding the `-6`/`-7` status
/// constants themselves. Every other status stays in [`Native`](Error::Native), which is where the
/// distinction is drawn: not "every code gets a variant" but "every code a caller needs to branch on
/// by name gets one".
#[derive(Debug, Clone)]
pub enum Error {
    /// Any native status not covered by a more specific variant below.
    Native { code: i32, message: String },
    /// The workbook is encrypted and no password was supplied (`XL_STATUS_PASSWORD_REQUIRED`).
    PasswordRequired { message: String },
    /// The supplied password did not match the workbook's verifier (`XL_STATUS_PASSWORD_INCORRECT`).
    PasswordIncorrect { message: String },
}

impl Error {
    /// Builds an [`Error`] from a raw `XL_*` status code and a detail message, routing
    /// `XL_STATUS_PASSWORD_REQUIRED`/`XL_STATUS_PASSWORD_INCORRECT` to their own variants and
    /// everything else to [`Native`](Error::Native). The one place that mapping happens, so every
    /// call site that turns a status code into an `Error` (here, in `arrow`/`writer`, and in
    /// `excelreader_derive`'s generated `append_tokens`) goes through it instead of re-deriving the
    /// mapping. `pub`, not `pub(crate)`: the derive macro expands into the crate that invokes it,
    /// not this one, so its generated code needs a public constructor to build a well-formed `Error`
    /// without reaching into a private field or duplicating this match.
    #[must_use]
    pub fn from_status(code: i32, message: String) -> Self {
        match code {
            crate::XL_STATUS_PASSWORD_REQUIRED => Error::PasswordRequired { message },
            crate::XL_STATUS_PASSWORD_INCORRECT => Error::PasswordIncorrect { message },
            _ => Error::Native { code, message },
        }
    }

    /// The `XL_*` status constant this error carries.
    #[must_use]
    pub fn code(&self) -> i32 {
        match self {
            Error::Native { code, .. } => *code,
            Error::PasswordRequired { .. } => crate::XL_STATUS_PASSWORD_REQUIRED,
            Error::PasswordIncorrect { .. } => crate::XL_STATUS_PASSWORD_INCORRECT,
        }
    }

    /// The detail from `xl_last_error_ptr` at the time of failure.
    #[must_use]
    pub fn message(&self) -> &str {
        match self {
            Error::Native { message, .. }
            | Error::PasswordRequired { message }
            | Error::PasswordIncorrect { message } => message,
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "ExcelReader error {}: {}", self.code(), self.message())
    }
}

impl std::error::Error for Error {}
