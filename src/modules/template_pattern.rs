use csv::StringRecord;
use std::error::Error;
use std::path::PathBuf;

pub struct TemplateStruct {
    pub csv_filepath: PathBuf,
    pub xlsx_filepath: PathBuf,
}

impl TemplateStruct {
    pub fn new(csv_filepath: PathBuf, xlsx_filepath: PathBuf) -> TemplateStruct {
        TemplateStruct {
            csv_filepath,
            xlsx_filepath,
        }
    }
}

pub trait TemplateManager {
    fn excute(&self) -> Result<(), Box<dyn Error>> {
        let records = self.get();
        self.set(records?)?;
        self.write()?;
        Ok(())
    }
    fn get(&self) -> Result<Vec<StringRecord>, Box<dyn Error>>;
    fn set(&self, records: Vec<StringRecord>) -> Result<(), Box<dyn Error>>;
    fn write(&self) -> Result<(), Box<dyn Error>>;
}
