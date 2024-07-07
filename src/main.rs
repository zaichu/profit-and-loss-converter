use clap::Parser;
use modules::dividend_list::lib::DividendListManager;
use modules::profit_and_loss::lib::ProfitAndLossManager;
use modules::settings::SETTINGS;
use modules::template_pattern::TemplateManager;
use std::error::Error;
use std::path::PathBuf;

mod modules;

#[derive(Parser)]
struct Args {
    #[clap(name = "CSVFILE")]
    csv_filepath: PathBuf,
    #[clap(name = "XLSXFILE")]
    xlsx_filepath: PathBuf,
}

enum FactoryID {
    ProfitAndLoss,
    DividendList,
}

fn create_factory(id: FactoryID, xlsx_filepath: PathBuf) -> Box<dyn TemplateManager> {
    match id {
        FactoryID::ProfitAndLoss => Box::new(ProfitAndLossManager::new(xlsx_filepath)),
        FactoryID::DividendList => Box::new(DividendListManager::new(xlsx_filepath)),
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    // CSVファイルとXLSXファイルのパスを取得する
    let csv_filepath = args.csv_filepath;
    let xlsx_filepath = args.xlsx_filepath;

    let factroy_id = csv_filepath
        .file_name()
        .and_then(|x| x.to_str())
        .map(|filename| {
            if filename.starts_with(&SETTINGS.prefix_profit_and_loss) {
                FactoryID::ProfitAndLoss
            } else if filename.starts_with(&SETTINGS.prefix_dividendlist) {
                FactoryID::DividendList
            } else {
                panic!("Invalid filename prefix.")
            }
        })
        .unwrap_or_else(|| panic!("Failed to extract filename."));

    // ファクトリからTemplateManagerを生成して実行する
    let factory = create_factory(factroy_id, xlsx_filepath);
    factory.execute(csv_filepath)?;

    Ok(())
}
