//! Round-trips through `xl_write_typed`: everything written here is read back with the same
//! crate's reader, so a layout mistake on either side shows up as a value mismatch rather than
//! as a file only Excel could judge.

use excelreader::{
    write_columns, write_columns_to, write_sheet, write_sheet_to, Column, ColumnData, Date,
    ErrorKind, ExcelMapper, Format, OpenOptions, Time, Timestamp, Workbook, WriteOptions,
};
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

fn xlsx() -> WriteOptions {
    WriteOptions::new().format(Format::Xlsx)
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
    let datas = [20454i32, 20455];
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
    write_columns(&path, &columns, &WriteOptions::new()).expect("write_columns must succeed");

    let workbook = Workbook::open(&path).expect("the written file must open");
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<WrittenRow>(1, 1)
        .expect("the written file must parse back");
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
fn write_columns_to_a_writer_round_trips() {
    let offsets = [0i32, 3, 6];
    let blob = b"umadoi";
    let inteiros = [1i64, 2];

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
    ];

    let mut bytes = Vec::new();
    write_columns_to(&mut bytes, &columns, &xlsx()).expect("write_columns_to must succeed");
    assert!(!bytes.is_empty());

    let workbook = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("the written bytes must open");
    #[derive(Default, Debug, ExcelMapper)]
    struct TwoColumnRow {
        #[excel(name = "texto")]
        texto: String,
        #[excel(name = "inteiro")]
        inteiro: i64,
    }
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<TwoColumnRow>(1, 1)
        .expect("the written bytes must parse back");
    assert_eq!(table.len(), 2);
    assert_eq!(table.get(0).expect("row 0").texto, "uma");
    assert_eq!(table.get(0).expect("row 0").inteiro, 1);
}

#[test]
fn writing_to_a_writer_needs_an_explicit_format() {
    let values = [1i64];
    let columns = [Column {
        name: Some("a"),
        data: ColumnData::I64(&values),
        validity: None,
    }];
    let error = write_columns_to(Vec::new(), &columns, &WriteOptions::new())
        .expect_err("a writer has no extension to infer a format from");
    assert_eq!(error.kind(), ErrorKind::InvalidArgument);
}

#[test]
fn write_columns_writes_nulls_from_the_validity_bitmap() {
    let nomes_offsets = [0i32, 3, 8, 8];
    let nomes = b"AnaBruno";
    let idades = [30i64, 41, 0];
    let idades_validity = [0b0000_0011u8];
    let pesos = [62.5f64, 0.0, 70.0];
    let pesos_validity = [0b0000_0101u8];
    let columns = [
        Column {
            name: Some("nome"),
            data: ColumnData::Str {
                offsets: &nomes_offsets,
                data: nomes,
            },
            validity: None,
        },
        Column {
            name: Some("idade"),
            data: ColumnData::I64(&idades),
            validity: Some(&idades_validity),
        },
        Column {
            name: Some("peso"),
            data: ColumnData::F64(&pesos),
            validity: Some(&pesos_validity),
        },
    ];

    #[derive(Default, Debug, ExcelMapper)]
    struct NullableRow {
        #[excel(name = "nome")]
        nome: String,
        #[excel(name = "idade")]
        idade: Option<i64>,
        #[excel(name = "peso")]
        peso: Option<f64>,
    }

    let path = temp_path("nullable.xlsx");
    write_columns(&path, &columns, &WriteOptions::new()).expect("write_columns must succeed");

    let workbook = Workbook::open(&path).expect("the written file must open");
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<NullableRow>(1, 1)
        .expect("the written file must parse back");
    let rows: Vec<(String, Option<i64>, Option<f64>)> = table
        .iter()
        .map(|row| (row.nome, row.idade, row.peso))
        .collect();
    assert_eq!(
        rows,
        vec![
            ("Ana".to_string(), Some(30), Some(62.5)),
            ("Bruno".to_string(), Some(41), None),
            (String::new(), None, Some(70.0)),
        ]
    );

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[test]
fn write_columns_rejects_inconsistent_input() {
    let two = [1i64, 2];
    let three = [1i64, 2, 3];
    let path = temp_path("rejected.xlsx");
    let options = WriteOptions::new();

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
    assert!(write_columns(&path, &mismatched, &options).is_err());

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
    assert!(write_columns(&path, &partial_header, &options).is_err());

    let short_bitmap = [Column {
        name: Some("a"),
        data: ColumnData::I64(&two),
        validity: Some(&[]),
    }];
    assert!(write_columns(&path, &short_bitmap, &options).is_err());

    let fine = [Column {
        name: Some("a"),
        data: ColumnData::I64(&two),
        validity: None,
    }];
    assert!(
        write_columns(temp_path("unknown.txt"), &fine, &options).is_err(),
        "an unrecognized extension with Format::Auto must be rejected"
    );

    assert!(write_columns(&path, &[], &options).is_err());
}

#[test]
fn write_columns_honors_the_sheet_name_and_infers_csv() {
    let valores = [1i64];
    let columns = [Column {
        name: Some("a"),
        data: ColumnData::I64(&valores),
        validity: None,
    }];

    let path = temp_path("named.xlsx");
    let options = WriteOptions::new().sheet_name("Dados");
    write_columns(&path, &columns, &options).expect("write must succeed");

    let workbook = Workbook::open(&path).expect("the written file must open");
    assert_eq!(
        workbook.sheet(0).expect("sheet 0").name().expect("name must read back"),
        "Dados"
    );
    drop(workbook);
    std::fs::remove_file(&path).ok();

    let csv = temp_path("out.csv");
    write_columns(&csv, &columns, &WriteOptions::new()).expect("a CSV write must succeed");
    let text = std::fs::read_to_string(&csv).expect("the .csv path must hold CSV text");
    assert_eq!(text.lines().collect::<Vec<_>>(), ["a", "1"]);
    std::fs::remove_file(&csv).ok();
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

fn derived_rows() -> Vec<DerivedRow> {
    vec![
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
    ]
}

#[test]
fn the_derive_round_trips_every_supported_field_type() {
    let rows = derived_rows();
    let path = temp_path("derived.xlsx");
    write_sheet(&path, &rows, &WriteOptions::new()).expect("write_sheet must succeed");

    let workbook = Workbook::open(&path).expect("the written file must open");
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<DerivedRow>(1, 1)
        .expect("the written file must parse back");
    assert_eq!(table.len(), 2);
    assert_eq!(table.get(0).expect("row 0"), rows[0]);
    assert_eq!(table.get(1).expect("row 1"), rows[1]);

    drop(table);
    std::fs::remove_file(&path).ok();
}

#[test]
fn write_sheet_to_a_writer_round_trips_the_derive() {
    let rows = derived_rows();
    let mut bytes = Vec::new();
    write_sheet_to(&mut bytes, &rows, &xlsx()).expect("write_sheet_to must succeed");

    let workbook = OpenOptions::new()
        .open_bytes(&bytes)
        .expect("the written bytes must open");
    let table = workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<DerivedRow>(1, 1)
        .expect("the written bytes must parse back");
    assert_eq!(table.len(), 2);
    assert_eq!(table.get(0).expect("row 0"), rows[0]);
    assert_eq!(table.get(1).expect("row 1"), rows[1]);
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
    write_sheet(&path, &rows, &WriteOptions::new()).expect("write_sheet must succeed");

    let workbook = Workbook::open(&path).expect("the written file must open");
    workbook
        .sheet(0)
        .expect("sheet 0")
        .parse::<AliasedRead>(1, 1)
        .expect("the header must carry the primary name, not the alias");

    std::fs::remove_file(&path).ok();
}
