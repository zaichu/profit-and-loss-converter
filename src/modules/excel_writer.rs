use super::profit_and_loss::ProfitAndLoss;
use std::error::Error;
use std::path::Path;
use umya_spreadsheet::*;

pub struct ExcelWriter;

impl ExcelWriter {
    const SHEET_NAME: &'static str = "株取引";
    const COLOR_ORANGE: &'static str = "FFF8CBAD";
    const COLOR_GREEN: &'static str = "FFC5E0B4";
    const COLOR_WHITE: &'static str = "FFFFFFFF";
    const HEADER: &'static [&'static str] = &[
        "約定日",
        "受渡日",
        "銘柄コード",
        "銘柄名",
        "口座",
        "数量[株]",
        "売却/決済単価[円]",
        "売却/決済額[円]",
        "平均取得価額[円]",
        "実現損益[円]",
        "源泉徴収税額",
        "損益",
    ];

    pub fn update_sheet(
        profit_and_loss: Vec<ProfitAndLoss>,
        xlsx_filepath: &Path,
    ) -> Result<(), Box<dyn Error>> {
        let mut book = reader::xlsx::read(xlsx_filepath)?;
        if book.get_sheet_by_name(Self::SHEET_NAME).is_some() {
            book.remove_sheet_by_name(Self::SHEET_NAME)?;
        }

        let new_sheet = book.new_sheet(Self::SHEET_NAME)?;
        Self::write_header(new_sheet);

        writer::xlsx::write(&book, xlsx_filepath)?;

        Ok(())
    }

    fn write_header(sheet: &mut Worksheet) {
        for (index, header) in Self::HEADER.iter().enumerate() {
            let cell = sheet.get_cell_mut((index as u32 + 2, 2));
            cell.set_value(header.to_string());
            Self::apply_style(cell, Self::COLOR_ORANGE);
        }
    }

    fn apply_style(cell: &mut umya_spreadsheet::Cell, color: &str) {
        let style = cell.get_style_mut();
        style.set_background_color(color);

        let borders = style.get_borders_mut();
        borders
            .get_bottom_border_mut()
            .set_border_style(Border::BORDER_THIN);
        borders
            .get_left_border_mut()
            .set_border_style(Border::BORDER_THIN);
        borders
            .get_right_border_mut()
            .set_border_style(Border::BORDER_THIN);
        borders
            .get_top_border_mut()
            .set_border_style(Border::BORDER_THIN);
    }
}
