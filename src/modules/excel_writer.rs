use super::{profit_and_loss::ProfitAndLoss, settings::SETTINGS};
use chrono::NaiveDate;
use std::collections::BTreeMap;
use std::error::Error;
use std::path::Path;
use umya_spreadsheet::{self, reader, writer, Border, Worksheet};

pub struct ExcelWriter;

impl ExcelWriter {
    const START_ROW: u32 = 2;
    const START_COL: u32 = 2;

    pub fn update_sheet(
        profit_and_loss_map: BTreeMap<NaiveDate, Vec<ProfitAndLoss>>,
        xlsx_filepath: &Path,
    ) -> Result<(), Box<dyn Error>> {
        let mut book = reader::xlsx::read(xlsx_filepath).expect("Failed to read xlsx");
        if book.get_sheet_by_name(&SETTINGS.sheet_name).is_some() {
            book.remove_sheet_by_name(&SETTINGS.sheet_name)?;
        }

        let mut new_sheet = book.new_sheet(&SETTINGS.sheet_name)?;
        Self::write_profit_and_loss(&mut new_sheet, profit_and_loss_map)?;
        Self::adjust_column_widths(&mut new_sheet)?;

        writer::xlsx::write(&book, xlsx_filepath)?;
        Ok(())
    }

    fn write_profit_and_loss(
        sheet: &mut Worksheet,
        profit_and_loss_map: BTreeMap<NaiveDate, Vec<ProfitAndLoss>>,
    ) -> Result<(), Box<dyn Error>> {
        Self::write_header(sheet);

        let mut row_index = Self::START_ROW;
        for (_, profit_and_loss) in profit_and_loss_map {
            let (specific_account_total, nisa_account_total) =
                Self::write_records(sheet, &mut row_index, profit_and_loss)?;
            row_index += 1;

            Self::write_footer(sheet, row_index, specific_account_total, nisa_account_total)?;
        }

        Ok(())
    }

    fn write_header(sheet: &mut Worksheet) {
        for (col_index, header) in ProfitAndLoss::HEADER.iter().enumerate() {
            let col_index = col_index as u32 + Self::START_COL;
            Self::write_cell(
                sheet,
                (col_index, Self::START_ROW),
                &(Some(header), None, None),
                Some(SETTINGS.colors.get("header_background").unwrap()),
            );
        }
    }

    fn write_records(
        sheet: &mut Worksheet,
        row_index: &mut u32,
        records: Vec<ProfitAndLoss>,
    ) -> Result<(i32, i32), Box<dyn Error>> {
        let mut specific_account_total = 0;
        let mut nisa_account_total = 0;

        for record in records {
            *row_index += 1;
            for (col_index, item) in record.get_profit_and_loss_struct_list().iter().enumerate() {
                let col_index = col_index as u32 + Self::START_COL;
                Self::write_cell(sheet, (col_index, *row_index), item, None);
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

        Ok((specific_account_total, nisa_account_total))
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
        for (col_index, item) in profit_and_loss
            .get_profit_and_loss_struct_list()
            .iter()
            .enumerate()
        {
            let col_index = col_index as u32 + Self::START_COL;
            Self::write_cell(
                sheet,
                (col_index, row_index),
                item,
                Some(SETTINGS.colors.get("footer_background").unwrap()),
            );
        }

        Ok(())
    }

    fn write_cell<T: ToString>(
        sheet: &mut Worksheet,
        coordinate: (u32, u32),
        (value, format, font_color): &(Option<T>, Option<&str>, Option<&str>),
        background_color: Option<&str>,
    ) {
        let cell = sheet.get_cell_mut(coordinate);
        // valueがNoneの場合は空文字列を設定
        if let Some(value) = value {
            cell.set_value(value.to_string());
        }

        if let Some(format) = format {
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(*format);
        }

        let style = cell.get_style_mut();
        if let Some(background_color) = background_color {
            style.set_background_color(background_color);
        }

        if let Some(font_color) = font_color {
            style.get_font_mut().get_color_mut().set_argb(*font_color);
        }

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

    fn adjust_column_widths(sheet: &mut Worksheet) -> Result<(), Box<dyn Error>> {
        for index in 2..=ProfitAndLoss::HEADER.len() as u32 {
            sheet
                .get_column_dimension_by_number_mut(&index)
                .set_width(15.0);
        }
        Ok(())
    }
}
