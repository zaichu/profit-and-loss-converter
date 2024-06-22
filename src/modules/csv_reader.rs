use super::profit_and_loss::ProfitAndLoss;
use chrono::NaiveDate;
use std::collections::BTreeMap;
use std::error::Error;
use std::path::Path;

pub struct CSVReader;

impl CSVReader {
    pub fn read_profit_and_loss(
        csv_filepath: &Path,
    ) -> Result<BTreeMap<NaiveDate, Vec<ProfitAndLoss>>, Box<dyn Error>> {
        let mut result = BTreeMap::new();
        let mut reader = csv::Reader::from_path(csv_filepath).expect("Failed to read csv");

        for record in reader.records() {
            let profit_and_loss = ProfitAndLoss::from_record(record?)?;
            if let Some(trade_date) = profit_and_loss.trade_date {
                result
                    .entry(trade_date)
                    .or_insert_with(Vec::new)
                    .push(profit_and_loss);
            }
        }

        Ok(result)
    }
}
