use super::profit_and_loss::ProfitAndLoss;
use std::error::Error;
use std::path::Path;

pub struct CSVReader;

impl CSVReader {
    pub fn read_profit_and_loss(csv_filepath: &Path) -> Result<Vec<ProfitAndLoss>, Box<dyn Error>> {
        let mut result = Vec::new();
        let mut reader = csv::Reader::from_path(csv_filepath)?;

        for record in reader.records() {
            result.push(ProfitAndLoss::new(record?)?);
        }

        Ok(result)
    }
}
