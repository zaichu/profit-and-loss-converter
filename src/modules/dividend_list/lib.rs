use super::{
    super::template_pattern::{TemplateManager, TemplateStruct},
    dividend_list::DividendList,
};
use crate::modules::{
    excel::{cell_style::CellStyle, coordinate::Coordinate, lib::ExcelAccessor},
    settings::SETTINGS,
};
use chrono::{Datelike, NaiveDate};
use csv::StringRecord;
use std::{cell::RefCell, collections::BTreeMap, error::Error, path::PathBuf};

pub struct DividendListManager {
    template_struct: TemplateStruct,
    dividend_list_map: RefCell<BTreeMap<NaiveDate, Vec<DividendList>>>,
}

impl DividendListManager {
    pub fn new(xlsx_filepath: PathBuf) -> Self {
        DividendListManager {
            template_struct: TemplateStruct::new(xlsx_filepath),
            dividend_list_map: RefCell::new(BTreeMap::new()),
        }
    }

    fn write_header(
        &self,
        excel_accessor: &mut ExcelAccessor,
        row_index: &mut u32,
    ) -> Result<(), Box<dyn Error>> {
        let header_list = DividendList::new().get_all_fields();

        for (col_index, (key, _)) in header_list.iter().enumerate() {
            let col_index = col_index as u32 + SETTINGS.start_col;
            let value = SETTINGS.headers.get(key).cloned();
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
        dividend_list: &[DividendList],
    ) -> Result<(i32, i32, i32), Box<dyn Error>> {
        let mut total_dividends_before_tax = 0; // 配当・分配金合計（税引前）[円/現地通貨]
        let mut total_taxes = 0; // 税額合計[円/現地通貨]
        let mut total_net_amount_received = 0; // 受取金額[円/現地通貨]

        for dividend in dividend_list {
            for (col_index, (field_name, value)) in dividend.get_all_fields().iter().enumerate() {
                let col_index = col_index as u32 + SETTINGS.start_col;
                let cell_style = &self.get_record_style(field_name, None);
                let coordinate_item = (col_index, *row_index).new_coordinate();
                excel_accessor.write_cell(coordinate_item, value, cell_style);
            }

            if let (Some(dividends_before_tax), Some(taxes), Some(net_amount_received)) = (
                dividend.dividends_before_tax,
                dividend.taxes,
                dividend.net_amount_received,
            ) {
                total_dividends_before_tax += dividends_before_tax;
                total_taxes += taxes;
                total_net_amount_received += net_amount_received;
            }

            *row_index += 1;
        }

        Ok((
            total_dividends_before_tax,
            total_taxes,
            total_net_amount_received,
        ))
    }

    fn get_record_style(&self, field_name: &str, background_color: Option<&String>) -> CellStyle {
        // background_color, font_format, font_color
        match field_name {
            "unit_price"// "単価[円/現地通貨]"
            | "dividends_before_tax" // "配当・分配金（税引前）[円/現地通貨]"
            | "taxes" // "税額[円/現地通貨]"
            | "net_amount_received" // "受取金額[円/現地通貨]"
            | "total_dividends_before_tax" // "配当・分配金合計（税引前）[円/現地通貨]"
            | "total_taxes" // "税額合計[円/現地通貨]"
            | "total_net_amount_received" // "受取金額合計[円/現地通貨]"
             => {
                CellStyle::new(background_color, SETTINGS.formats.get("yen"), None)
            }
            _ => CellStyle::new(background_color, None, None),
        }
    }

    fn write_footer(
        &self,
        excel_accessor: &mut ExcelAccessor,
        row_index: &mut u32,
        total: (i32, i32, i32),
    ) -> Result<(), Box<dyn Error>> {
        let dividend_list = DividendList::new_total_dividend_list(total);
        for (col_index, (field_name, value)) in dividend_list.get_all_fields().iter().enumerate() {
            let coordinate_item =
                (col_index as u32 + SETTINGS.start_col, *row_index).new_coordinate();
            let cell_style =
                &self.get_record_style(field_name, SETTINGS.colors.get("footer_background"));
            excel_accessor.write_cell(coordinate_item, value, cell_style);
        }

        *row_index += 1;

        Ok(())
    }
}

impl TemplateManager for DividendListManager {
    fn set(&self, records: Vec<StringRecord>) -> Result<(), Box<dyn Error>> {
        for record in records {
            let dividend = DividendList::from_record(record)?;
            if let Some(settlement_date) = dividend.settlement_date {
                let date =
                    NaiveDate::from_ymd_opt(settlement_date.year(), settlement_date.month(), 1)
                        .unwrap();
                self.dividend_list_map
                    .borrow_mut()
                    .entry(date)
                    .or_insert_with(Vec::new)
                    .push(dividend);
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

        for (_, dividend_list) in self.dividend_list_map.borrow_mut().iter_mut() {
            // 取引履歴書き込み
            let total = self.write_records(&mut excel_accessor, &mut row_index, dividend_list)?;

            self.write_footer(&mut excel_accessor, &mut row_index, total)?;
        }

        let len = DividendList::new().get_all_fields().len() as u32;
        excel_accessor.adjust_column_widths(len)?;
        excel_accessor.save_book()?;

        Ok(())
    }
}
