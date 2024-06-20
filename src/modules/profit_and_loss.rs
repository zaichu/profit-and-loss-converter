use chrono::NaiveDate;
use csv::StringRecord;
use std::error::Error;

#[derive(Debug, Clone)]
pub struct ProfitAndLoss {
    pub trade_date: NaiveDate,      // 約定日
    pub settlement_date: NaiveDate, // 受渡日
    pub securities_code: String,    // 銘柄コード
    pub company_name: String,       // 銘柄名
    pub account: String,            // 口座
    pub shares: i32,                // 数量[株]
    pub asked_price: f64,           // 売却/決済単価[円]
    pub settlement_amount: i32,     // 売却/決済額[円]
    pub purchase_price: f64,        // 平均取得価額[円]
    pub profit_and_loss: i32,       // 実現損益[円]
}

impl ProfitAndLoss {
    pub fn new(record: StringRecord) -> Result<Self, Box<dyn Error>> {
        Ok(ProfitAndLoss {
            trade_date: Self::parse_date("trade_date", record.get(0))?,
            settlement_date: Self::parse_date("settlement_date", record.get(1))?,
            securities_code: Self::parse_string("securities_code", record.get(2))?,
            company_name: Self::parse_string("company_name", record.get(3))?,
            account: Self::parse_string("account", record.get(4))?,
            shares: Self::parse_int("shares", record.get(7))?,
            asked_price: Self::parse_float("asked_price", record.get(8))?,
            settlement_amount: Self::parse_int("settlement_amount", record.get(9))?,
            purchase_price: Self::parse_float("purchase_price", record.get(10))?,
            profit_and_loss: Self::parse_int("profit_and_loss", record.get(11))?,
        })
    }

    pub fn get_profit_and_loss_list(&self) -> [String; 10] {
        [
            self.trade_date.to_string(),
            self.settlement_date.to_string(),
            self.securities_code.clone(),
            self.company_name.clone(),
            self.account.clone(),
            self.shares.to_string(),
            self.asked_price.to_string(),
            self.settlement_amount.to_string(),
            self.purchase_price.to_string(),
            self.profit_and_loss.to_string(),
        ]
    }

    fn parse_date(field_name: &str, date_str: Option<&str>) -> Result<NaiveDate, Box<dyn Error>> {
        let date_str = date_str.ok_or_else(|| format!("Missing {field_name} field"))?;
        NaiveDate::parse_from_str(date_str.replace("/", "-").as_str(), "%Y-%m-%d")
            .map_err(|e| format!("Failed to parse date '{date_str}': {e}").into())
    }

    fn parse_int(field_name: &str, num_str: Option<&str>) -> Result<i32, Box<dyn Error>> {
        let num_str = num_str.ok_or_else(|| format!("Missing {field_name} field"))?;
        num_str
            .replace(",", "")
            .parse::<i32>()
            .map_err(|e| format!("Failed to parse integer '{num_str}': {e}").into())
    }

    fn parse_float(field_name: &str, num_str: Option<&str>) -> Result<f64, Box<dyn Error>> {
        let num_str = num_str.ok_or_else(|| format!("Missing {field_name} field"))?;
        num_str
            .replace(",", "")
            .parse::<f64>()
            .map_err(|e| format!("Failed to parse float '{num_str}': {e}").into())
    }

    fn parse_string(field_name: &str, value: Option<&str>) -> Result<String, Box<dyn Error>> {
        let value = value.ok_or_else(|| format!("Missing {field_name} field"))?;
        Ok(value.to_string())
    }
}
