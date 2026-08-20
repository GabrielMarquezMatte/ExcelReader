use excelreader::workbook::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Nome")]
    nome: Vec<u8>,
}

fn main() {}
