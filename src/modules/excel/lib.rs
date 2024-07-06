use crate::modules::excel::{cell_style::CellStyle, coordinate::CoordinateItem};
use crate::modules::settings::SETTINGS;
use std::cell::RefCell;
use std::error::Error;
use std::path::PathBuf;
use umya_spreadsheet::{self, reader, writer, Border, Spreadsheet};

pub struct ExcelAccessor {
    book: RefCell<Spreadsheet>,
    sheet_title: String,
    xlsx_filepath: PathBuf,
}

impl ExcelAccessor {
    pub fn read_book(sheet_title: &str, xlsx_filepath: &PathBuf) -> Result<Self, Box<dyn Error>> {
        let mut book = reader::xlsx::read(xlsx_filepath).expect("Failed to read xlsx");
        if book.get_sheet_by_name(sheet_title).is_some() {
            book.remove_sheet_by_name(sheet_title)?;
        }
        _ = book.new_sheet(sheet_title)?;
        Ok(ExcelAccessor {
            book: RefCell::new(book),
            sheet_title: sheet_title.to_string(),
            xlsx_filepath: xlsx_filepath.clone(),
        })
    }

    pub fn write_cell(
        &mut self,
        coordinate: CoordinateItem,
        value: &Option<String>,
        cell_style: &CellStyle,
    ) {
        if let Some(sheet) = self
            .book
            .borrow_mut()
            .get_sheet_by_name_mut(self.sheet_title.as_str())
        {
            let cell = sheet.get_cell_mut((coordinate.col, coordinate.row));
            if let Some(value) = value {
                cell.set_value(value);
            }

            if let Some(format) = cell_style.font_format.clone() {
                cell.get_style_mut()
                    .get_number_format_mut()
                    .set_format_code(format);
            }

            let style = cell.get_style_mut();
            if let Some(background_color) = cell_style.background_color.clone() {
                style.set_background_color(background_color);
            }

            if let Some(font_color) = cell_style.font_color.clone() {
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
    }

    pub fn adjust_column_widths(&mut self, width: f64, len: u32) -> Result<(), Box<dyn Error>> {
        for index in SETTINGS.start_col..=SETTINGS.start_col + len {
            if let Some(sheet) = self
                .book
                .borrow_mut()
                .get_sheet_by_name_mut(self.sheet_title.as_str())
            {
                sheet
                    .get_column_dimension_by_number_mut(&index)
                    .set_width(width);
            }
        }
        Ok(())
    }

    pub fn save_book(&self) -> Result<(), Box<dyn Error>> {
        writer::xlsx::write(&self.book.borrow_mut(), &self.xlsx_filepath)?;
        Ok(())
    }
}
