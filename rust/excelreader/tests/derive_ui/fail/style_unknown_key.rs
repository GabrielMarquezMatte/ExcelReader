use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Valor", style(bold, colour = 0xFF0000))]
    valor: f64,
}

fn main() {}
