//! Adapted from the task brief's pseudocode to this crate's actual `Workbook::open_with`/
//! `sheet_count` signatures (`open_with` takes a `&str` path, an explicit format, and
//! `Option<&OpenOptions>`; `sheet_count` returns `Result<i32, Error>`, matching every other
//! `Workbook` accessor) - the brief's snippet assumed a simpler, two-argument `open_with` and an
//! infallible `sheet_count` that this crate does not have. The test intent (open with a password,
//! report `PasswordRequired`/`PasswordIncorrect`, and survive a password sourced from a temporary)
//! is unchanged.

use excelreader::workbook::Workbook;
use excelreader::{Error, OpenOptions, XL_FORMAT_AUTO};

fn fixture(name: &str) -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/ExcelReader.Tests/data/encrypted")
        .join(name)
        .to_str()
        .expect("fixture path must be UTF-8")
        .to_string()
}

#[test]
fn opens_encrypted_workbook_with_password() {
    let options = OpenOptions::new().password("hunter2");
    let book = Workbook::open_with(
        &fixture("agile-aes256-sha512.xlsx"),
        XL_FORMAT_AUTO,
        Some(&options),
    )
    .unwrap();
    assert!(book.sheet_count().unwrap() > 0);
}

#[test]
fn reports_password_required_without_password() {
    let err = Workbook::open_with(&fixture("agile-aes256-sha512.xlsx"), XL_FORMAT_AUTO, None)
        .unwrap_err();
    assert!(matches!(err, Error::PasswordRequired { .. }), "got {err:?}");
}

#[test]
fn reports_password_incorrect_with_wrong_password() {
    let options = OpenOptions::new().password("wrong");
    let err = Workbook::open_with(
        &fixture("agile-aes256-sha512.xlsx"),
        XL_FORMAT_AUTO,
        Some(&options),
    )
    .unwrap_err();
    assert!(
        matches!(err, Error::PasswordIncorrect { .. }),
        "got {err:?}"
    );
}

// The regression this task exists to prevent: the raw struct's password pointer must not dangle
// when the password came from a temporary.
#[test]
fn password_survives_a_temporary_source() {
    let options = OpenOptions::new().password(String::from("hunter2"));
    let book = Workbook::open_with(
        &fixture("agile-aes256-sha512.xlsx"),
        XL_FORMAT_AUTO,
        Some(&options),
    )
    .unwrap();
    assert!(book.sheet_count().unwrap() > 0);
}
