use chrono::NaiveDate;
use std::path::Path;

#[derive(Debug, Clone)]
pub struct ProfitAndLoss {
    pub trade_date: NaiveDate,      // 約定日
    pub settlement_date: NaiveDate, // 受渡日
    pub securities_code: String,    // 銘柄コード
    pub company_name: String,       // 銘柄名
    pub account: String,            // 口座
    pub shares: i32,                // 数量[株]
    pub asked_price: i32,           // 売却/決済単価[円]
    pub settlement_amount: i32,     // 売却/決済額[円]
    pub purchase_price: i32,        // 平均取得価額[円]
    pub profit_and_loss: i32,       // 実現損益[円]
}

pub fn read_csv(path: &Path) -> Result<(), csv::Error> {
    let mut reader = csv::Reader::from_path(path)?;
    for result in reader.records() {
        let data = result?;
        println!("{data:?}");
    }

    Ok(())
}
