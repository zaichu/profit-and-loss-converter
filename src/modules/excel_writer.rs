use super::profit_and_loss::ProfitAndLoss;
use chrono::NaiveDate;
use std::collections::BTreeMap;
use std::error::Error;
use std::path::Path;
use umya_spreadsheet::{self, reader, writer, Border, Cell, Worksheet};

pub struct ExcelWriter;

impl ExcelWriter {
    const SHEET_NAME: &'static str = "株取引";
    const TAX_RATE: f64 = 0.20315;
    const COLOR_ORANGE: &'static str = "FFF8CBAD";
    const COLOR_GREEN: &'static str = "FFC5E0B4";
    const COLOR_WHITE: &'static str = "FF000000";
    const START_ROW: u32 = 2;
    const START_COL: u32 = 2;
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
        profit_and_loss_map: BTreeMap<NaiveDate, Vec<ProfitAndLoss>>,
        xlsx_filepath: &Path,
    ) -> Result<(), Box<dyn Error>> {
        let mut book = reader::xlsx::read(xlsx_filepath)?;
        if book.get_sheet_by_name(Self::SHEET_NAME).is_some() {
            book.remove_sheet_by_name(Self::SHEET_NAME)?;
        }

        let mut new_sheet = book.new_sheet(Self::SHEET_NAME)?;

        Self::write_profit_and_loss(&mut new_sheet, profit_and_loss_map)?;

        for index in 2..=Self::HEADER.len() {
            let col = index as u32;
            new_sheet
                .get_column_dimension_by_number_mut(&col)
                .set_width(15.0);
        }

        writer::xlsx::write(&book, xlsx_filepath)?;

        Ok(())
    }

    fn write_header(sheet: &mut Worksheet) {
        for (col_index, header) in Self::HEADER.iter().enumerate() {
            let col_index = col_index as u32 + Self::START_COL;
            Self::write_value(
                sheet,
                (col_index, Self::START_ROW),
                header,
                None,
                Some(Self::COLOR_ORANGE),
            );
        }
    }

    fn write_footer(sheet: &mut Worksheet, row_index: u32, total: i32) {
        for col_index in 0..9 {
            let col_index = col_index as u32 + Self::START_COL;
            Self::write_value(
                sheet,
                (col_index, row_index),
                "",
                None,
                Some(Self::COLOR_GREEN),
            );
        }

        // 実現損益[円]
        Self::write_value(
            sheet,
            (11, row_index),
            total,
            Some(ProfitAndLoss::YEN_FORMAT),
            Some(Self::COLOR_GREEN),
        );

        let tax: i32 = if total < 0 {
            0
        } else {
            (total as f64 * Self::TAX_RATE) as i32
        };

        // 源泉徴収税額
        Self::write_value(
            sheet,
            (12, row_index),
            tax,
            Some(ProfitAndLoss::YEN_FORMAT),
            Some(Self::COLOR_GREEN),
        );

        // 損益
        Self::write_value(
            sheet,
            (13, row_index),
            total - tax,
            Some(ProfitAndLoss::YEN_FORMAT),
            Some(Self::COLOR_GREEN),
        );
    }

    fn write_profit_and_loss(
        sheet: &mut Worksheet,
        profit_and_loss_map: BTreeMap<NaiveDate, Vec<ProfitAndLoss>>,
    ) -> Result<(), Box<dyn Error>> {
        Self::write_header(sheet);

        let mut row_index = Self::START_ROW;
        for (_, profit_and_loss) in profit_and_loss_map {
            let mut total = 0;
            for record in profit_and_loss {
                row_index += 1;
                for (col_index, (value, format)) in
                    record.get_profit_and_loss_list().iter().enumerate()
                {
                    let col_index = col_index as u32 + Self::START_COL;
                    Self::write_value(sheet, (col_index, row_index), value, *format, None);
                }

                let col_index = record.get_profit_and_loss_list().len() as u32;
                Self::write_value(sheet, (col_index + 2, row_index), "-", None, None);
                Self::write_value(sheet, (col_index + 3, row_index), "-", None, None);

                total += record.realized_profit_and_loss;
            }
            row_index += 1;
            Self::write_footer(sheet, row_index, total);
        }

        Ok(())
    }

    fn write_value<T: ToString>(
        sheet: &mut Worksheet,
        coordinate: (u32, u32),
        value: T,
        format: Option<&str>,
        color: Option<&str>,
    ) {
        let cell = sheet.get_cell_mut(coordinate);
        cell.set_value(value.to_string());

        if let Some(format) = format {
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format);
        }

        if let Some(color) = color {
            Self::apply_style(cell, color);
        } else {
            Self::apply_style(cell, Self::COLOR_WHITE);
        }
    }

    fn apply_style(cell: &mut Cell, color: &str) {
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
