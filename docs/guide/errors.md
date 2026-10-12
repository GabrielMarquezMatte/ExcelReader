# Errors across languages

The .NET library throws exceptions. The C ABI turns each failure into an `XL_*` status code, and the
Python, Rust and C++ bindings turn that code into their own error type. This page maps one onto the
other, so a failure seen in one language can be recognized in another.

## Status codes

Every `xl_*` function returns one of these. After a failure, `xl_last_error_ptr` returns the message
for the calling thread.

| Code | Constant | Meaning |
|---|---|---|
| 0 | `XL_OK` | Success. |
| -1 | `XL_EOF` | No more rows or batches. Not an error: the bindings end iteration on it. |
| -2 | `XL_BUFFER_TOO_SMALL` | The caller's buffer cannot hold the result. The bindings grow the buffer and retry, so it never reaches user code. |
| -3 | `XL_INVALID_HANDLE` | The handle is null, already closed, or was never opened. |
| -4 | `XL_INVALID_ARGUMENT` | An argument failed validation before any work started: a null pointer, a bad length, malformed options. |
| -5 | `XL_ERROR` | Anything else. The message is the message of the .NET exception behind it. |
| -6 | `XL_STATUS_PASSWORD_REQUIRED` | The workbook is encrypted and no password was supplied. |
| -7 | `XL_STATUS_PASSWORD_INCORRECT` | The supplied password did not match the workbook's verifier. |

## By failure

| Failure | C# | Status | Python | Rust | C++ |
|---|---|---|---|---|---|
| Encrypted, no password | `ExcelEncryptionException`, `Reason == PasswordRequired` | -6 | `PasswordRequiredError` | `Error::PasswordRequired` | `xl::Error`, `code == XL_STATUS_PASSWORD_REQUIRED` |
| Wrong password | `ExcelEncryptionException`, `Reason == PasswordIncorrect` | -7 | `PasswordIncorrectError` | `Error::PasswordIncorrect` | `xl::Error`, `code == XL_STATUS_PASSWORD_INCORRECT` |
| Unsupported encryption, or failed integrity check | `ExcelEncryptionException`, `Reason` is `UnsupportedScheme` or `IntegrityFailure` | -5 | `ExcelReaderError` | `Error::Native` | `xl::Error`, `code == XL_ERROR` |
| Corrupt or malformed file | `InvalidDataException` | -5 | `ExcelReaderError` | `Error::Native` | `xl::Error`, `code == XL_ERROR` |
| Feature the reader does not support | `NotSupportedException` | -5 | `ExcelReaderError` | `Error::Native` | `xl::Error`, `code == XL_ERROR` |
| File exceeds a configured resource limit | `ExcelLimitExceededException` | -5 | `ExcelReaderError` | `Error::Native` | `xl::Error`, `code == XL_ERROR` |
| Workbook already closed | `ObjectDisposedException` | -3 | `ExcelReaderError` | `Error::Native` | `xl::Error`, `code == XL_INVALID_HANDLE` |
| Invalid argument | `ArgumentException` and its subclasses | -4, or -5 when the check happens past the ABI boundary | `ExcelReaderError` | `Error::Native` | `xl::Error` |

`PasswordRequiredError` and `PasswordIncorrectError` derive from `ExcelReaderError`, so
`except ExcelReaderError` catches every native failure in Python. Rust's `Error::code()` returns the
status for every variant. The C++ header throws nothing: every entry point returns
`std::expected<T, xl::Error>`.

Only the two password cases are worth retrying with different input. The rest fail the same way on
every attempt.

Each binding also checks some arguments itself, before calling the ABI, and reports those with its
own native types: Python raises `ValueError`, `TypeError` or `IndexError`; Rust panics on an
out-of-range index (see the crate README's "Bounds and panics"); C++ returns `xl::Error` with
`code == XL_INVALID_ARGUMENT`.
