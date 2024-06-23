use super::{profit_and_loss::ProfitAndLoss, settings::SETTINGS};
use chrono::NaiveDate;
use std::collections::BTreeMap;
use std::error::Error;
use std::path::Path;
use umya_spreadsheet::{self, reader, writer, Border, Worksheet};

pub struct ExcelWriter;

impl ExcelWriter {
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
        Self::write_header(sheet, ProfitAndLoss::new().unwrap());

        let mut row_index = SETTINGS.start_row;
        for (_, profit_and_loss) in profit_and_loss_map {
            let (specific_account_total, nisa_account_total) =
                Self::write_records(sheet, &mut row_index, profit_and_loss)?;
            row_index += 1;

            Self::write_footer(sheet, row_index, specific_account_total, nisa_account_total)?;
        }

        Ok(())
    }

    fn write_header(sheet: &mut Worksheet, header: ProfitAndLoss) {
        for (col_index, (key, _)) in header.get_profit_and_loss_struct_list().iter().enumerate() {
            let col_index = col_index as u32 + SETTINGS.start_col;
            let value = SETTINGS.headers.get(*key).cloned();
            let header_background_color = SETTINGS.colors.get("header_background");
            Self::write_cell(
                sheet,
                (col_index, SETTINGS.start_row),
                (value, None, header_background_color, None),
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
            for (col_index, (key, value)) in
                record.get_profit_and_loss_struct_list().iter().enumerate()
            {
                let col_index = col_index as u32 + SETTINGS.start_col;
                Self::write_cell(
                    sheet,
                    (col_index, *row_index),
                    Self::get_record_style((key, value.clone())),
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
        for (col_index, (key, value)) in profit_and_loss
            .get_profit_and_loss_struct_list()
            .iter()
            .enumerate()
        {
            let col_index = col_index as u32 + SETTINGS.start_col;
            Self::write_cell(
                sheet,
                (col_index, row_index),
                Self::get_footter_style((key, value.clone())),
            );
        }

        Ok(())
    }

    fn write_cell(
        sheet: &mut Worksheet,
        coordinate: (u32, u32),
        (value, format, background_color, font_color): (
            Option<String>,
            Option<&String>,
            Option<&String>,
            Option<&String>,
        ),
    ) {
        let cell = sheet.get_cell_mut(coordinate);
        if let Some(value) = value {
            cell.set_value(value.to_string());
        }

        if let Some(format) = format {
            cell.get_style_mut()
                .get_number_format_mut()
                .set_format_code(format);
        }

        let style = cell.get_style_mut();
        if let Some(background_color) = background_color {
            style.set_background_color(background_color);
        }

        if let Some(font_color) = font_color {
            style.get_font_mut().get_color_mut().set_argb(font_color);
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
        for index in SETTINGS.start_col..=SETTINGS.start_col + SETTINGS.headers.len() as u32 {
            sheet
                .get_column_dimension_by_number_mut(&index)
                .set_width(15.0);
        }
        Ok(())
    }

    fn get_record_style(
        (key, value): (&str, Option<String>),
    ) -> (
        Option<String>,
        Option<&String>,
        Option<&String>,
        Option<&String>,
    ) {
        // (value, format, background_color, font_color)
        let yen_decimal_format = SETTINGS.formats.get("yen_decimal");
        let yen_format = SETTINGS.formats.get("yen");
        let realized_loss_font_color = SETTINGS.colors.get("realized_loss_font");

        match (key, value.clone()) {
            // "売却/決済単価[円]",　"平均取得価額[円]"
            (stringify!(asked_price), Some(_)) | (stringify!(purchase_price), Some(_)) => {
                (value, yen_decimal_format, None, None)
            }
            // "売却/決済額[円]", "実現損益[円]"
            (stringify!(proceeds), Some(_value))
            | (stringify!(realized_profit_and_loss), Some(_value)) => {
                if value.clone().unwrap().starts_with('-') {
                    (value, yen_format, None, realized_loss_font_color)
                } else {
                    (value, yen_format, None, None)
                }
            }
            _ => (value, None, None, None),
        }
    }

    fn get_footter_style(
        (key, value): (&str, Option<String>),
    ) -> (
        Option<String>,
        Option<&String>,
        Option<&String>,
        Option<&String>,
    ) {
        // (value, format, background_color, font_color)
        let yen_format = SETTINGS.formats.get("yen");
        let background_color = SETTINGS.colors.get("footer_background").clone();
        let realized_loss_font_color = SETTINGS.colors.get("realized_loss_font");

        match (key, value.clone()) {
            (stringify!(trade_date), Some(_value))
            | (stringify!(withholding_tax), Some(_value))
            | (stringify!(profit_and_loss), Some(_value)) => {
                if _value.starts_with('-') {
                    (
                        value,
                        yen_format,
                        background_color,
                        realized_loss_font_color,
                    )
                } else {
                    (value, yen_format, background_color, None)
                }
            }
            _ => (value, None, background_color, None),
        }
    }
}
