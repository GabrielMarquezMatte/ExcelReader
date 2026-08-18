use excelreader::workbook::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Nome")]
    nome: u32,
}

fn main() {}
