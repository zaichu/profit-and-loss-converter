use chrono::NaiveDate;
use csv::StringRecord;
use std::error::Error;

#[derive(Debug, Clone)]
pub struct ProfitAndLoss {
    pub trade_date: Option<NaiveDate>,
    pub settlement_date: Option<NaiveDate>,
    pub security_code: Option<String>,
    pub security_name: Option<String>,
    pub account: Option<String>,
    pub shares: Option<i32>,
    pub asked_price: Option<f64>,
    pub proceeds: Option<i32>,
    pub purchase_price: Option<f64>,
    pub realized_profit_and_loss: Option<i32>,
    pub total_realized_profit_and_loss: Option<i32>,
    pub withholding_tax: Option<u32>,
    pub profit_and_loss: Option<i32>,
}

impl ProfitAndLoss {
    const YEN_DECIMAL_FORMAT: &'static str = "\"¥\"#,##0.00;\"¥\"-#,##0.00";
    const YEN_FORMAT: &'static str = "\"¥\"#,##0;\"¥\"-#,##0";
    const TAX_RATE: f64 = 0.20315;
    pub const HEADER: &'static [&'static str] = &[
        "約定日",
        "受渡日",
        "銘柄コード",
        "銘柄名",
        "口座",
        "数量[株]",
        "売却/決済単価[円]",
        "売却/決済額[円]",
        "平均取得価額[円]",
        "実現損益[円]",
        "合計実現損益[円]",
        "源泉徴収税額",
        "損益",
    ];

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

    pub fn with_total_realized_profit_and_loss(
        specific_account_total: i32,
        nisa_account_total: i32,
    ) -> Result<Self, Box<dyn Error>> {
        let withholding_tax = if specific_account_total < 0 {
            0
        } else {
            (specific_account_total as f64 * Self::TAX_RATE) as u32
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

    pub fn get_profit_and_loss_struct_list(
        &self,
    ) -> [(Option<String>, Option<&'static str>); Self::HEADER.len()] {
        [
            (self.trade_date.map(|d| d.to_string()), None),
            (self.settlement_date.map(|d| d.to_string()), None),
            (self.security_code.clone(), None),
            (self.security_name.clone(), None),
            (self.account.clone(), None),
            (self.shares.map(|s| s.to_string()), None),
            (
                self.asked_price.map(|p| p.to_string()),
                Some(Self::YEN_DECIMAL_FORMAT),
            ),
            (
                self.proceeds.map(|p| p.to_string()),
                Some(Self::YEN_DECIMAL_FORMAT),
            ),
            (
                self.purchase_price.map(|p| p.to_string()),
                Some(Self::YEN_DECIMAL_FORMAT),
            ),
            (
                self.realized_profit_and_loss.map(|p| p.to_string()),
                Some(Self::YEN_DECIMAL_FORMAT),
            ),
            (
                self.total_realized_profit_and_loss.map(|p| p.to_string()),
                Some(Self::YEN_FORMAT),
            ),
            (
                self.withholding_tax.map(|p| p.to_string()),
                Some(Self::YEN_FORMAT),
            ),
            (
                self.profit_and_loss.map(|p| p.to_string()),
                Some(Self::YEN_FORMAT),
            ),
        ]
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
