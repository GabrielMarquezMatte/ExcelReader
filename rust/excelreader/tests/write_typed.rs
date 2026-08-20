//! Round-trips through `xl_write_typed`: everything written here is read back with the same
//! crate's reader, so a layout mistake on either side shows up as a value mismatch rather than
//! as a file only Excel could judge.

use excelreader::workbook::{parse_sheet, ExcelMapper, Workbook};
use excelreader::writer::{Column, ColumnData, ExcelWriter, OwnedColumn, OwnedColumnData};
use excelreader::{Date, Time, Timestamp, XL_FORMAT_CSV, XL_FORMAT_XLSX};
use std::path::PathBuf;
use std::sync::atomic::{AtomicU32, Ordering};

static COUNTER: AtomicU32 = AtomicU32::new(0);

/// A unique path under the system temp directory. `tempfile` would do this too, but a dev
/// dependency to build one path is more machinery than the job needs.
fn temp_path(name: &str) -> PathBuf {
    let mut path = std::env::temp_dir();
    path.push(format!(
        "excelreader-rs-{}-{}-{name}",
        std::process::id(),
        COUNTER.fetch_add(1, Ordering::Relaxed)
    ));
    path
}

#[derive(Default, Debug, ExcelMapper)]
struct WrittenRow {
    #[excel(name = "texto")]
    texto: String,
    #[excel(name = "inteiro")]
    inteiro: i64,
    #[excel(name = "numero")]
    numero: f64,
    #[excel(name = "ativo")]
    ativo: bool,
    #[excel(name = "data")]
    data: Date,
    #[excel(name = "hora")]
    hora: Time,
    #[excel(name = "instante")]
    instante: Timestamp,
}

#[test]
fn write_columns_round_trips_every_column_type() {
    let offsets = [0i32, 3, 6];
    let blob = b"umadoi";
    let inteiros = [1i64, 2];
    let numeros = [0.5f64, 1.5];
    let ativos = [1u8, 0];
    let datas = [20454i32, 20455]; // 2026-01-01, 2026-01-02
    let horas = [3_600_000_000i64, 7_200_000_000];
    let instantes = [1_767_225_600_000_000i64, 1_767_312_000_000_000];

    let columns = [
        Column {
            name: Some("texto"),
            data: ColumnData::Str {
                offsets: &offsets,
                data: blob,
            },
            validity: None,
        },
        Column {
            name: Some("inteiro"),
            data: ColumnData::I64(&inteiros),
            validity: None,
        },
        Column {
            name: Some("numero"),
            data: ColumnData::F64(&numeros),
            validity: None,
        },
        Column {
            name: Some("ativo"),
            data: ColumnData::Bool(&ativos),
            validity: None,
        },
        Column {
            name: Some("data"),
            data: ColumnData::Date(&datas),
            validity: None,
        },
        Column {
            name: Some("hora"),
            data: ColumnData::Time(&horas),
            validity: None,
        },
        Column {
            name: Some("instante"),
            data: ColumnData::Timestamp(&instantes),
            validity: None,
        },
    ];

    let path = temp_path("columns.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &columns, None)
        .expect("write_columns must succeed");

    let mut workbook = Workbook::open(target).expect("the written file must open");
    let table =
        parse_sheet::<WrittenRow>(&mut workbook, 1).expect("the written file must parse back");
    assert_eq!(table.len(), 2);

    let first = table.get(0).expect("row 0 must exist");
    assert_eq!(first.texto, "uma");
    assert_eq!(first.inteiro, 1);
    assert_eq!(first.numero, 0.5);
    assert!(first.ativo);
    assert_eq!(first.data, Date::new(20454));
    assert_eq!(first.hora, Time::new(3_600_000_000));
    assert_eq!(first.instante, Timestamp::new(1_767_225_600_000_000));

    let second = table.get(1).expect("row 1 must exist");
    assert_eq!(second.texto, "doi");
    assert!(!second.ativo);

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[test]
fn write_columns_writes_nulls_from_the_validity_bitmap() {
    let valores = [10i64, 0, 30];
    // LSB-first: bits 0 and 2 set, bit 1 clear - row 1 is null.
    let validity = [0b0000_0101u8];
    let columns = [Column {
        name: Some("quantidade"),
        data: ColumnData::I64(&valores),
        validity: Some(&validity),
    }];

    #[derive(Default, Debug, ExcelMapper)]
    struct NullableRow {
        #[excel(name = "quantidade")]
        quantidade: Option<i64>,
    }

    let path = temp_path("nullable.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &columns, None)
        .expect("write_columns must succeed");

    let mut workbook = Workbook::open(target).expect("the written file must open");
    let table =
        parse_sheet::<NullableRow>(&mut workbook, 1).expect("the written file must parse back");
    assert_eq!(table.get(0).unwrap().quantidade, Some(10));
    assert_eq!(table.get(1).unwrap().quantidade, None);
    assert_eq!(table.get(2).unwrap().quantidade, Some(30));

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[test]
fn write_columns_rejects_inconsistent_input() {
    let two = [1i64, 2];
    let three = [1i64, 2, 3];
    let path = temp_path("rejected.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");

    let mismatched = [
        Column {
            name: Some("a"),
            data: ColumnData::I64(&two),
            validity: None,
        },
        Column {
            name: Some("b"),
            data: ColumnData::I64(&three),
            validity: None,
        },
    ];
    assert!(excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &mismatched, None).is_err());

    let partial_header = [
        Column {
            name: Some("a"),
            data: ColumnData::I64(&two),
            validity: None,
        },
        Column {
            name: None,
            data: ColumnData::I64(&two),
            validity: None,
        },
    ];
    assert!(
        excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &partial_header, None).is_err()
    );

    // Two rows need one byte of bitmap; hand it an empty slice.
    let short_bitmap = [Column {
        name: Some("a"),
        data: ColumnData::I64(&two),
        validity: Some(&[]),
    }];
    assert!(
        excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &short_bitmap, None).is_err()
    );

    let fine = [Column {
        name: Some("a"),
        data: ColumnData::I64(&two),
        validity: None,
    }];
    assert!(
        excelreader::writer::write_columns(target, excelreader::XL_FORMAT_AUTO, &fine, None)
            .is_err(),
        "XL_FORMAT_AUTO must be rejected: a new file has no signature bytes to sniff"
    );

    assert!(excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &[], None).is_err());
}

#[test]
fn write_columns_honors_the_sheet_name_and_csv_format() {
    let valores = [1i64];
    let columns = [Column {
        name: Some("a"),
        data: ColumnData::I64(&valores),
        validity: None,
    }];

    let path = temp_path("named.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    let options = excelreader::WriteOptions::new().sheet_name("Dados");
    excelreader::writer::write_columns(target, XL_FORMAT_XLSX, &columns, Some(&options))
        .expect("write must succeed");

    let workbook = Workbook::open(target).expect("the written file must open");
    assert_eq!(workbook.sheet_name().expect("name must read back"), "Dados");
    drop(workbook);
    std::fs::remove_file(&path).ok();

    let csv = temp_path("out.csv");
    let csv_target = csv.to_str().expect("temp path must be UTF-8");
    assert_eq!(
        excelreader::writer::format_from_path(csv_target),
        XL_FORMAT_CSV
    );
    excelreader::writer::write_columns(csv_target, XL_FORMAT_CSV, &columns, None)
        .expect("a CSV write must succeed");
    std::fs::remove_file(&csv).ok();
}

struct ManualRow {
    nome: String,
    idade: i64,
    peso: Option<f64>,
}

impl ExcelWriter for ManualRow {
    fn to_columns(rows: &[Self]) -> Result<Vec<OwnedColumn>, excelreader::Error> {
        let n = rows.len();
        let mut nome_offsets: Vec<i32> = Vec::with_capacity(n + 1);
        nome_offsets.push(0);
        let mut nome_data: Vec<u8> = Vec::new();
        let mut idade: Vec<i64> = Vec::with_capacity(n);
        let mut peso: Vec<f64> = Vec::with_capacity(n);
        let mut peso_validity: Vec<u8> = vec![0; n.div_ceil(8)];

        for (row, r) in rows.iter().enumerate() {
            excelreader::writer::push_str(&mut nome_offsets, &mut nome_data, r.nome.as_str())?;
            idade.push(r.idade);
            match &r.peso {
                Some(value) => {
                    excelreader::writer::set_valid(&mut peso_validity, row);
                    peso.push(*value);
                }
                None => peso.push(f64::default()),
            }
        }

        Ok(vec![
            OwnedColumn {
                name: Some("nome"),
                data: OwnedColumnData::Str {
                    offsets: nome_offsets,
                    data: nome_data,
                },
                validity: None,
            },
            OwnedColumn {
                name: Some("idade"),
                data: OwnedColumnData::I64(idade),
                validity: None,
            },
            OwnedColumn {
                name: Some("peso"),
                data: OwnedColumnData::F64(peso),
                validity: Some(peso_validity),
            },
        ])
    }
}

#[derive(Default, Debug, ExcelMapper)]
struct ManualRowRead {
    #[excel(name = "nome")]
    nome: String,
    #[excel(name = "idade")]
    idade: i64,
    #[excel(name = "peso")]
    peso: Option<f64>,
}

#[test]
fn write_sheet_round_trips_a_hand_written_excel_writer() {
    let rows = vec![
        ManualRow {
            nome: "Ana".to_string(),
            idade: 30,
            peso: Some(62.5),
        },
        ManualRow {
            nome: "Bruno".to_string(),
            idade: 41,
            peso: None,
        },
    ];

    let path = temp_path("manual.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    excelreader::writer::write_sheet(target, XL_FORMAT_XLSX, &rows, None)
        .expect("write_sheet must succeed");

    let mut workbook = Workbook::open(target).expect("the written file must open");
    let table =
        parse_sheet::<ManualRowRead>(&mut workbook, 1).expect("the written file must parse back");
    assert_eq!(table.len(), 2);

    let first = table.get(0).expect("row 0 must exist");
    assert_eq!(first.nome, "Ana");
    assert_eq!(first.idade, 30);
    assert_eq!(first.peso, Some(62.5));

    let second = table.get(1).expect("row 1 must exist");
    assert_eq!(second.nome, "Bruno");
    assert_eq!(second.peso, None);

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[derive(Default, Debug, PartialEq, ExcelMapper)]
struct DerivedRow {
    #[excel(name = "texto")]
    texto: String,
    #[excel(name = "inteiro", alias = "int")]
    inteiro: i32,
    #[excel(name = "numero")]
    numero: f32,
    #[excel(name = "ativo")]
    ativo: bool,
    #[excel(name = "data")]
    data: Date,
    #[excel(name = "hora")]
    hora: Time,
    #[excel(name = "instante")]
    instante: Timestamp,
    #[excel(name = "opcional")]
    opcional: Option<i64>,
}

#[test]
fn the_derive_round_trips_every_supported_field_type() {
    let rows = vec![
        DerivedRow {
            texto: "uma".to_string(),
            inteiro: 1,
            numero: 0.5,
            ativo: true,
            data: Date::new(20454),
            hora: Time::new(3_600_000_000),
            instante: Timestamp::new(1_767_225_600_000_000),
            opcional: Some(7),
        },
        DerivedRow {
            texto: "duas".to_string(),
            inteiro: 2,
            numero: 1.5,
            ativo: false,
            data: Date::new(20455),
            hora: Time::new(7_200_000_000),
            instante: Timestamp::new(1_767_312_000_000_000),
            opcional: None,
        },
    ];

    let path = temp_path("derived.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    excelreader::writer::write_sheet(target, XL_FORMAT_XLSX, &rows, None)
        .expect("write_sheet must succeed");

    let mut workbook = Workbook::open(target).expect("the written file must open");
    let table =
        parse_sheet::<DerivedRow>(&mut workbook, 1).expect("the written file must parse back");
    assert_eq!(table.len(), 2);
    assert_eq!(table.get(0).expect("row 0"), rows[0]);
    assert_eq!(table.get(1).expect("row 1"), rows[1]);

    drop(table);
    std::fs::remove_file(&path).ok();
}

/// The write side uses only the PRIMARY name. An alias that reached a write spec would be
/// rejected by the ABI ("must have exactly one name"), so this asserts the header actually
/// written is `inteiro`, never `int`.
#[test]
fn the_derive_writes_only_the_primary_column_name() {
    #[derive(Default, Debug, ExcelMapper)]
    struct AliasedRead {
        #[excel(name = "inteiro")]
        inteiro: i32,
    }

    let rows = vec![DerivedRow::default()];
    let path = temp_path("aliased.xlsx");
    let target = path.to_str().expect("temp path must be UTF-8");
    excelreader::writer::write_sheet(target, XL_FORMAT_XLSX, &rows, None)
        .expect("write_sheet must succeed");

    let mut workbook = Workbook::open(target).expect("the written file must open");
    parse_sheet::<AliasedRead>(&mut workbook, 1)
        .expect("the header must carry the primary name, not the alias");

    std::fs::remove_file(&path).ok();
}
