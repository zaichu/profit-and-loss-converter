use super::profit_and_loss::ProfitAndLoss;
use chrono::NaiveDate;

use std::collections::BTreeMap;
use std::error::Error;
use std::path::Path;
use umya_spreadsheet::{self, reader, writer, Border, Cell, Worksheet};

pub struct ExcelWriter;

impl ExcelWriter {
    const SHEET_NAME: &'static str = "株取引";

    const COLOR_ORANGE: &'static str = "FFF8CBAD";
    const COLOR_GREEN: &'static str = "FFC5E0B4";
    const COLOR_WHITE: &'static str = "FF000000";
    const START_ROW: u32 = 2;
    const START_COL: u32 = 2;

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

        for index in 2..=ProfitAndLoss::HEADER.len() {
            let col = index as u32;
            new_sheet
                .get_column_dimension_by_number_mut(&col)
                .set_width(15.0);
        }

        writer::xlsx::write(&book, xlsx_filepath)?;

        Ok(())
    }

    fn write_header(sheet: &mut Worksheet) {
        for (col_index, header) in ProfitAndLoss::HEADER.iter().enumerate() {
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

    fn write_footer(
        sheet: &mut Worksheet,
        row_index: u32,
        specific_account_total: i32,
        nisa_account_total: i32,
    ) -> Result<(), Box<dyn Error>> {
        let profit_and_loss = ProfitAndLoss::with_total_realized_profit_and_loss(
            specific_account_total,
            nisa_account_total,
        )?;

        for (col_index, (value, format)) in profit_and_loss
            .get_profit_and_loss_struct_list()
            .iter()
            .enumerate()
        {
            let col_index = col_index as u32 + Self::START_COL;

            Self::write_value(
                sheet,
                (col_index, row_index),
                value.as_deref().unwrap_or(""),
                *format,
                Some(Self::COLOR_GREEN),
            );
        }

        Ok(())
    }

    fn write_profit_and_loss(
        sheet: &mut Worksheet,
        profit_and_loss_map: BTreeMap<NaiveDate, Vec<ProfitAndLoss>>,
    ) -> Result<(), Box<dyn Error>> {
        Self::write_header(sheet);

        let mut row_index = Self::START_ROW;
        for (_, profit_and_loss) in profit_and_loss_map {
            let mut specific_account_total = 0;
            let mut nisa_account_total = 0;
            for record in profit_and_loss {
                row_index += 1;

                for (col_index, (value, format)) in
                    record.get_profit_and_loss_struct_list().iter().enumerate()
                {
                    let col_index = col_index as u32 + Self::START_COL;
                    Self::write_value(
                        sheet,
                        (col_index, row_index),
                        value.as_deref().unwrap_or("-"),
                        *format,
                        None,
                    );
                }

                if let (Some(account), Some(realized_profit_and_loss)) =
                    (record.account, record.realized_profit_and_loss)
                {
                    if account.contains("特定") {
                        specific_account_total += realized_profit_and_loss;
                    } else {
                        nisa_account_total += realized_profit_and_loss;
                    }
                }
            }
            row_index += 1;

            let _ =
                Self::write_footer(sheet, row_index, specific_account_total, nisa_account_total);
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

        Self::apply_style(cell, color.as_deref().unwrap_or(Self::COLOR_WHITE));
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
