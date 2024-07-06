#[derive(Clone, Debug)]
pub struct CellStyle {
    pub background_color: Option<String>,
    pub font_format: Option<String>,
    pub font_color: Option<String>,
}

impl CellStyle {
    pub fn new(
        background_color: Option<&String>,
        font_format: Option<&String>,
        font_color: Option<&String>,
    ) -> Self {
        CellStyle {
            background_color: background_color.cloned(),
            font_format: font_format.cloned(),
            font_color: font_color.cloned(),
        }
    }
}
