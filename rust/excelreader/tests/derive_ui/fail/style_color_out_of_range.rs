use excelreader::ExcelMapper;

#[derive(Default, ExcelMapper)]
struct Row {
    #[excel(name = "Valor", style(font_color = 0x1000000))]
    valor: f64,
}

fn main() {}
