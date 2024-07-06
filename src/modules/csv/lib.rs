use csv::StringRecord;
use std::error::Error;
use std::path::Path;

pub struct CSVAccessor;

impl CSVAccessor {
    pub fn read(filepath: &Path) -> Result<Vec<StringRecord>, Box<dyn Error>> {
        let mut result: Vec<StringRecord> = Vec::new();
        let mut reader = csv::Reader::from_path(filepath).expect("Failed to read csv");
        for record in reader.records() {
            result.push(record?);
        }
        Ok(result)
    }
}
