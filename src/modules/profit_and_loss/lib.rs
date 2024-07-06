use super::super::template_pattern::{TemplateManager, TemplateStruct};
use super::profit_and_loss::ProfitAndLoss;
use crate::modules::csv::lib::CSVAccessor;
use crate::modules::excel::{cell_style::CellStyle, coordinate::Coordinate, lib::ExcelAccessor};
use crate::modules::settings::SETTINGS;
use chrono::NaiveDate;
use csv::StringRecord;
use std::cell::RefCell;
use std::collections::BTreeMap;
use std::error::Error;
use std::path::PathBuf;

pub struct ProfitAndLossManager {
    template_struct: TemplateStruct,
    profit_and_loss_map: RefCell<BTreeMap<NaiveDate, Vec<ProfitAndLoss>>>,
}

impl ProfitAndLossManager {
    pub fn new(csv_filepath: PathBuf, xlsx_filepath: PathBuf) -> Self {
        ProfitAndLossManager {
            template_struct: TemplateStruct::new(csv_filepath, xlsx_filepath),
            profit_and_loss_map: RefCell::new(BTreeMap::new()),
        }
    }

    fn write_header(
        &self,
        excel_accessor: &mut ExcelAccessor,
        row_index: &mut u32,
    ) -> Result<(), Box<dyn Error>> {
        let header_list = ProfitAndLoss::new()?.get_all_fields();

        for (col_index, (key, _)) in header_list.iter().enumerate() {
            let col_index = col_index as u32 + SETTINGS.start_col;
            let value = SETTINGS.headers.get(key).map(|x| x.clone());
            let background_color = SETTINGS.colors.get("header_background");
            let coordinate_item = (col_index, *row_index).new_coordinate();
            excel_accessor.write_cell(
                coordinate_item,
                &value,
                &CellStyle::new(background_color, None, None),
            );
        }
        *row_index += 1;
        Ok(())
    }

    fn write_records(
        &self,
        excel_accessor: &mut ExcelAccessor,
        row_index: &mut u32,
        profit_and_loss_list: Vec<ProfitAndLoss>,
    ) -> Result<(i32, i32), Box<dyn Error>> {
        let mut specific_account_total = 0;
        let mut nisa_account_total = 0;

        for profit_and_loss in profit_and_loss_list {
            for (col_index, (field_name, value)) in
                profit_and_loss.get_all_fields().iter().enumerate()
            {
                let col_index = col_index as u32 + SETTINGS.start_col;
                let cell_style = &self.get_record_style(field_name, value);
                let coordinate_item = (col_index, *row_index).new_coordinate();
                excel_accessor.write_cell(coordinate_item, value, cell_style);
            }

            if let (Some(account), Some(realized_profit_and_loss)) = (
                profit_and_loss.account,
                profit_and_loss.realized_profit_and_loss,
            ) {
                if account.contains("特定") {
                    specific_account_total += realized_profit_and_loss;
                } else {
                    nisa_account_total += realized_profit_and_loss;
                }
            }

            *row_index += 1;
        }

        Ok((specific_account_total, nisa_account_total))
    }

    fn get_record_style(&self, field_name: &str, value: &Option<String>) -> CellStyle {
        let yen_decimal_format = SETTINGS.formats.get("yen_decimal");
        let yen_format = SETTINGS.formats.get("yen");
        let realized_loss_font_color = SETTINGS.colors.get("realized_loss_font");

        // background_color, font_format, font_color
        match (field_name, value) {
            // "売却/決済単価[円]",　"平均取得価額[円]"
            ("asked_price", Some(_)) | ("purchase_price", Some(_)) => {
                CellStyle::new(None, yen_decimal_format, None)
            }
            // "売却/決済額[円]", "実現損益[円]"
            ("proceeds", Some(value)) | ("realized_profit_and_loss", Some(value)) => {
                if value.starts_with('-') {
                    CellStyle::new(None, yen_format, realized_loss_font_color)
                } else {
                    CellStyle::new(None, yen_format, None)
                }
            }
            _ => CellStyle::new(None, None, None),
        }
    }

    fn write_footer(
        &self,
        excel_accessor: &mut ExcelAccessor,
        row_index: &mut u32,
        specific_account_total: i32,
        nisa_account_total: i32,
    ) -> Result<(), Box<dyn Error>> {
        let profit_and_loss = ProfitAndLoss::new_total_realized_profit_and_loss(
            specific_account_total,
            nisa_account_total,
        )?;
        for (col_index, (field_name, value)) in profit_and_loss.get_all_fields().iter().enumerate()
        {
            let coordinate_item =
                (col_index as u32 + SETTINGS.start_col, *row_index).new_coordinate();
            let cell_style = &self.get_footer_style(field_name, value);
            excel_accessor.write_cell(coordinate_item, value, cell_style);
        }

        *row_index += 1;

        Ok(())
    }

    fn get_footer_style(&self, field_name: &str, value: &Option<String>) -> CellStyle {
        let yen_format = SETTINGS.formats.get("yen");
        let background_color = SETTINGS.colors.get("footer_background").clone();
        let realized_loss_font_color = SETTINGS.colors.get("realized_loss_font");

        // background_color, font_format, font_color
        match (field_name, value) {
            ("total_realized_profit_and_loss", Some(value))
            | ("withholding_tax", Some(value))
            | ("profit_and_loss", Some(value)) => {
                if value.starts_with('-') {
                    CellStyle::new(background_color, yen_format, realized_loss_font_color)
                } else {
                    CellStyle::new(background_color, yen_format, None)
                }
            }
            _ => CellStyle::new(background_color, None, None),
        }
    }
}

impl TemplateManager for ProfitAndLossManager {
    fn get(&self) -> Result<Vec<StringRecord>, Box<dyn Error>> {
        CSVAccessor::read(&self.template_struct.csv_filepath)
    }

    fn set(&self, records: Vec<StringRecord>) -> Result<(), Box<dyn Error>> {
        for record in records {
            let profit_and_loss = ProfitAndLoss::from_record(record)?;
            if let Some(trade_date) = profit_and_loss.trade_date {
                self.profit_and_loss_map
                    .borrow_mut()
                    .entry(trade_date)
                    .or_insert_with(Vec::new)
                    .push(profit_and_loss);
            }
        }

        Ok(())
    }

    fn write(&self) -> Result<(), Box<dyn Error>> {
        let mut excel_accessor =
            ExcelAccessor::read_book(&SETTINGS.sheet_title, &self.template_struct.xlsx_filepath)?;

        // ヘッダー書き込み
        let mut row_index = SETTINGS.start_row;
        self.write_header(&mut excel_accessor, &mut row_index)?;

        for (_, profit_and_loss) in self.profit_and_loss_map.borrow_mut().iter() {
            // 取引履歴書き込み
            let (specific_account_total, nisa_account_total) =
                self.write_records(&mut excel_accessor, &mut row_index, profit_and_loss.clone())?;

            self.write_footer(
                &mut excel_accessor,
                &mut row_index,
                specific_account_total,
                nisa_account_total,
            )?
        }

        let len = ProfitAndLoss::new()?.get_all_fields().len() as u32;
        excel_accessor.adjust_column_widths(16.0, len)?;
        excel_accessor.save_book()?;
        Ok(())
    }
}