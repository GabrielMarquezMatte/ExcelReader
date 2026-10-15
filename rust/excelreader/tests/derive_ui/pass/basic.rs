use excelreader::ExcelMapper;
use excelreader::ColumnType;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Nome")]
    nome: String,
    #[excel(name = "Idade")]
    idade: i64,
    #[excel(name = "Peso")]
    peso: f64,
    #[excel(name = "Ativo")]
    ativo: Option<bool>,
}

fn main() {
    let bindings = Row::bindings();
    assert_eq!(bindings.len(), 4);
    assert_eq!(bindings[0].names[0], "Nome");
    assert_eq!(bindings[0].column_type, ColumnType::String);
    assert_eq!(bindings[1].names[0], "Idade");
    assert_eq!(bindings[1].column_type, ColumnType::I64);
    assert_eq!(bindings[2].names[0], "Peso");
    assert_eq!(bindings[2].column_type, ColumnType::F64);
    assert_eq!(bindings[3].names[0], "Ativo");
    assert_eq!(bindings[3].column_type, ColumnType::Bool);
}
