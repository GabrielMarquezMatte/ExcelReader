//! Adapted from the task brief's pseudocode to this crate's actual `Workbook::open_with`/
//! `sheet_count` signatures (`open_with` takes a `&str` path, an explicit format, and
//! `Option<&OpenOptions>`; `sheet_count` returns `Result<i32, Error>`, matching every other
//! `Workbook` accessor) - the brief's snippet assumed a simpler, two-argument `open_with` and an
//! infallible `sheet_count` that this crate does not have. The test intent (open with a password,
//! report `PasswordRequired`/`PasswordIncorrect`, and survive a password sourced from a temporary)
//! is unchanged.

use excelreader::workbook::Workbook;
use excelreader::writer::{encrypt_package, write_columns, Column, ColumnData};
use excelreader::{Error, OpenOptions, XL_FORMAT_AUTO, XL_FORMAT_XLSX};
use std::path::PathBuf;
use std::sync::atomic::{AtomicU32, Ordering};

fn fixture(name: &str) -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/ExcelReader.Tests/data/encrypted")
        .join(name)
        .to_str()
        .expect("fixture path must be UTF-8")
        .to_string()
}

static COUNTER: AtomicU32 = AtomicU32::new(0);

/// A unique path under the system temp directory - same helper as write_typed.rs's, duplicated
/// rather than shared since these are separate integration test binaries with no common crate to
/// put it in.
fn temp_path(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!(
        "excelreader-rs-{}-{}-{name}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    path
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
fn encrypt_package_produces_a_file_openable_with_the_same_password() {
    let inteiros = [1i64, 2];
    let columns = [Column {
        name: Some("inteiro"),
        data: ColumnData::I64(&inteiros),
        validity: None,
    }];

    let plain_path = temp_path("encrypt-plain.xlsx");
    let encrypted_path = temp_path("encrypt-cipher.xlsx");
    let plain_target = plain_path.to_str().expect("temp path must be UTF-8");
    let encrypted_target = encrypted_path.to_str().expect("temp path must be UTF-8");

    write_columns(plain_target, XL_FORMAT_XLSX, &columns, None)
        .expect("write_columns must succeed");
    encrypt_package(plain_target, encrypted_target, "hunter2").expect("encrypt_package must succeed");

    let options = OpenOptions::new().password("hunter2");
    let book = Workbook::open_with(encrypted_target, XL_FORMAT_AUTO, Some(&options))
        .expect("the encrypted file must open with the right password");
    assert!(book.sheet_count().unwrap() > 0);

    let err = Workbook::open_with(encrypted_target, XL_FORMAT_AUTO, None).unwrap_err();
    assert!(matches!(err, Error::PasswordRequired { .. }), "got {err:?}");

    let _ = std::fs::remove_file(&plain_path);
    let _ = std::fs::remove_file(&encrypted_path);
}

#[test]
fn encrypt_package_rejects_an_empty_password() {
    let inteiros = [1i64];
    let columns = [Column {
        name: Some("inteiro"),
        data: ColumnData::I64(&inteiros),
        validity: None,
    }];

    let plain_path = temp_path("encrypt-empty-pw.xlsx");
    let plain_target = plain_path.to_str().expect("temp path must be UTF-8");
    write_columns(plain_target, XL_FORMAT_XLSX, &columns, None)
        .expect("write_columns must succeed");

    let destination = temp_path("encrypt-empty-pw-out.xlsx");
    let err = encrypt_package(plain_target, destination.to_str().unwrap(), "").unwrap_err();
    assert_eq!(err.code(), excelreader::XL_INVALID_ARGUMENT, "got {err:?}");

    let _ = std::fs::remove_file(&plain_path);
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
