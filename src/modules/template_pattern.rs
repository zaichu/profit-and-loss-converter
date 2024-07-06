use crate::modules::csv::lib::CSVAccessor;
use csv::StringRecord;
use std::error::Error;
use std::path::PathBuf;

pub struct TemplateStruct {
    pub xlsx_filepath: PathBuf,
}

impl TemplateStruct {
    pub fn new(xlsx_filepath: PathBuf) -> TemplateStruct {
        TemplateStruct { xlsx_filepath }
    }
}

pub trait TemplateManager {
    fn excute(&self, csv_filepath: PathBuf) -> Result<(), Box<dyn Error>> {
        let records = self.get(csv_filepath);
        self.set(records?)?;
        self.write()?;
        Ok(())
    }
    fn get(&self, csv_filepath: PathBuf) -> Result<Vec<StringRecord>, Box<dyn Error>> {
        CSVAccessor::read(&csv_filepath)
    }
    fn set(&self, records: Vec<StringRecord>) -> Result<(), Box<dyn Error>>;
    fn write(&self) -> Result<(), Box<dyn Error>>;
}
