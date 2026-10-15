use excelreader::{
    encrypt_package, write_columns, Column, ColumnData, ErrorKind, OpenOptions, Workbook,
    WriteOptions,
};
use std::path::PathBuf;
use std::sync::atomic::{AtomicU32, Ordering};

fn fixture(name: &str) -> PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../tests/ExcelReader.Tests/data/encrypted")
        .join(name)
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
    let book = OpenOptions::new()
        .password("hunter2")
        .open(fixture("agile-aes256-sha512.xlsx"))
        .unwrap();
    assert!(!book.sheets().unwrap().is_empty());
}

#[test]
fn encrypt_package_produces_a_file_openable_with_the_same_password() {
    let inteiros = [1i64, 2];
    let columns = [Column::new(Some("inteiro"), ColumnData::I64(&inteiros))];

    let plain_path = temp_path("encrypt-plain.xlsx");
    let encrypted_path = temp_path("encrypt-cipher.xlsx");

    write_columns(&plain_path, &columns, &WriteOptions::new()).expect("write_columns must succeed");
    encrypt_package(&plain_path, &encrypted_path, "hunter2").expect("encrypt_package must succeed");

    let book = OpenOptions::new()
        .password("hunter2")
        .open(&encrypted_path)
        .expect("the encrypted file must open with the right password");
    assert!(!book.sheets().unwrap().is_empty());

    let err = Workbook::open(&encrypted_path).unwrap_err();
    assert_eq!(err.kind(), ErrorKind::PasswordRequired, "got {err:?}");

    let _ = std::fs::remove_file(&plain_path);
    let _ = std::fs::remove_file(&encrypted_path);
}

#[test]
fn encrypt_package_rejects_an_empty_password() {
    let inteiros = [1i64];
    let columns = [Column::new(Some("inteiro"), ColumnData::I64(&inteiros))];

    let plain_path = temp_path("encrypt-empty-pw.xlsx");
    write_columns(&plain_path, &columns, &WriteOptions::new()).expect("write_columns must succeed");

    let destination = temp_path("encrypt-empty-pw-out.xlsx");
    let err = encrypt_package(&plain_path, &destination, "").unwrap_err();
    assert_eq!(err.kind(), ErrorKind::InvalidArgument, "got {err:?}");

    let _ = std::fs::remove_file(&plain_path);
}

#[test]
fn reports_password_required_without_password() {
    let err = Workbook::open(fixture("agile-aes256-sha512.xlsx")).unwrap_err();
    assert_eq!(err.kind(), ErrorKind::PasswordRequired, "got {err:?}");
}

#[test]
fn reports_password_incorrect_with_wrong_password() {
    let err = OpenOptions::new()
        .password("wrong")
        .open(fixture("agile-aes256-sha512.xlsx"))
        .unwrap_err();
    assert_eq!(err.kind(), ErrorKind::PasswordIncorrect, "got {err:?}");
}

#[test]
fn password_survives_a_temporary_source() {
    let options = OpenOptions::new().password(String::from("hunter2"));
    let book = options.open(fixture("agile-aes256-sha512.xlsx")).unwrap();
    assert!(!book.sheets().unwrap().is_empty());
}
