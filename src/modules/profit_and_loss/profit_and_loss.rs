use super::super::settings::SETTINGS;
use chrono::NaiveDate;
use csv::StringRecord;
use std::error::Error;

#[derive(Debug, Clone)]
pub struct ProfitAndLoss {
    pub trade_date: Option<NaiveDate>,               // 約定日
    pub settlement_date: Option<NaiveDate>,          // 受渡日
    pub security_code: Option<String>,               // 銘柄コード
    pub security_name: Option<String>,               // 銘柄名
    pub account: Option<String>,                     // 口座
    pub shares: Option<i32>,                         // 数量[株]
    pub asked_price: Option<f64>,                    // 売却/決済単価[円]
    pub proceeds: Option<i32>,                       // 売却/決済額[円]
    pub purchase_price: Option<f64>,                 // 平均取得価額[円]
    pub realized_profit_and_loss: Option<i32>,       // 実現損益[円]
    pub total_realized_profit_and_loss: Option<i32>, // 合計実現損益[円]
    pub withholding_tax: Option<u32>,                // 源泉徴収税額
    pub profit_and_loss: Option<i32>,                // 損益
}

impl ProfitAndLoss {
    pub fn new() -> Result<Self, &'static str> {
        Ok(ProfitAndLoss {
            trade_date: None,
            settlement_date: None,
            security_code: None,
            security_name: None,
            account: None,
            shares: None,
            asked_price: None,
            proceeds: None,
            purchase_price: None,
            realized_profit_and_loss: None,
            total_realized_profit_and_loss: None,
            withholding_tax: None,
            profit_and_loss: None,
        })
    }

    pub fn from_record(record: StringRecord) -> Result<Self, Box<dyn Error>> {
        Ok(ProfitAndLoss {
            trade_date: Self::parse_date(record.get(0))?,
            settlement_date: Self::parse_date(record.get(1))?,
            security_code: Self::parse_string(record.get(2))?,
            security_name: Self::parse_string(record.get(3))?,
            account: Self::parse_string(record.get(4))?,
            shares: Self::parse_int(record.get(7))?,
            asked_price: Self::parse_float(record.get(8))?,
            proceeds: Self::parse_int(record.get(9))?,
            purchase_price: Self::parse_float(record.get(10))?,
            realized_profit_and_loss: Self::parse_int(record.get(11))?,
            total_realized_profit_and_loss: None,
            withholding_tax: None,
            profit_and_loss: None,
        })
    }

    pub fn get_all_fields(&self) -> Vec<(String, Option<String>)> {
        vec![
            (
                "trade_date".to_string(),
                self.trade_date.map(|d| d.to_string()),
            ),
            (
                "settlement_date".to_string(),
                self.settlement_date.map(|d| d.to_string()),
            ),
            ("security_code".to_string(), self.security_code.clone()),
            ("security_name".to_string(), self.security_name.clone()),
            ("account".to_string(), self.account.clone()),
            ("shares".to_string(), self.shares.map(|s| s.to_string())),
            (
                "asked_price".to_string(),
                self.asked_price.map(|p| p.to_string()),
            ),
            ("proceeds".to_string(), self.proceeds.map(|p| p.to_string())),
            (
                "purchase_price".to_string(),
                self.purchase_price.map(|p| p.to_string()),
            ),
            (
                "realized_profit_and_loss".to_string(),
                self.realized_profit_and_loss.map(|p| p.to_string()),
            ),
            (
                "total_realized_profit_and_loss".to_string(),
                self.total_realized_profit_and_loss.map(|p| p.to_string()),
            ),
            (
                "withholding_tax".to_string(),
                self.withholding_tax.map(|p| p.to_string()),
            ),
            (
                "profit_and_loss".to_string(),
                self.profit_and_loss.map(|p| p.to_string()),
            ),
        ]
    }

    pub fn new_total_realized_profit_and_loss(
        (specific_account_total, nisa_account_total): (i32, i32),
    ) -> Result<Self, Box<dyn Error>> {
        let withholding_tax = if specific_account_total < 0 {
            0
        } else {
            (specific_account_total as f64 * SETTINGS.tax_rate) as u32
        };
        let total = specific_account_total + nisa_account_total;

        Ok(ProfitAndLoss {
            trade_date: None,
            settlement_date: None,
            security_code: None,
            security_name: None,
            account: None,
            shares: None,
            asked_price: None,
            proceeds: None,
            purchase_price: None,
            realized_profit_and_loss: None,
            total_realized_profit_and_loss: Some(total),
            withholding_tax: Some(withholding_tax),
            profit_and_loss: Some(total - withholding_tax as i32),
        })
    }

    fn parse_date(date_str: Option<&str>) -> Result<Option<NaiveDate>, Box<dyn Error>> {
        date_str.map_or(Ok(None), |s| {
            NaiveDate::parse_from_str(&s.replace("/", "-"), "%Y-%m-%d")
                .map(Some)
                .map_err(|e| format!("Failed to parse date '{s}': {e}").into())
        })
    }

    fn parse_int(num_str: Option<&str>) -> Result<Option<i32>, Box<dyn Error>> {
        num_str.map_or(Ok(None), |s| {
            s.replace(",", "")
                .parse::<i32>()
                .map(Some)
                .map_err(|e| format!("Failed to parse integer '{s}': {e}").into())
        })
    }

    fn parse_float(num_str: Option<&str>) -> Result<Option<f64>, Box<dyn Error>> {
        num_str.map_or(Ok(None), |s| {
            s.replace(",", "")
                .parse::<f64>()
                .map(Some)
                .map_err(|e| format!("Failed to parse float '{s}': {e}").into())
        })
    }

    fn parse_string(value: Option<&str>) -> Result<Option<String>, Box<dyn Error>> {
        Ok(value.map(|s| s.to_string()))
    }
}
