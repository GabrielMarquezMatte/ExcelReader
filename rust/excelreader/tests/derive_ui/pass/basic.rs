use excelreader::workbook::ExcelMapper;
use excelreader::{XL_T_BOOL, XL_T_F64, XL_T_I64, XL_T_STRING};

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
    assert_eq!(bindings[0].name, "Nome");
    assert_eq!(bindings[0].xl_type, XL_T_STRING);
    assert_eq!(bindings[1].name, "Idade");
    assert_eq!(bindings[1].xl_type, XL_T_I64);
    assert_eq!(bindings[2].name, "Peso");
    assert_eq!(bindings[2].xl_type, XL_T_F64);
    assert_eq!(bindings[3].name, "Ativo");
    assert_eq!(bindings[3].xl_type, XL_T_BOOL);
}
