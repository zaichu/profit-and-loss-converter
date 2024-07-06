use clap::Parser;
use modules::profit_and_loss::lib::ProfitAndLossManager;
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
}

fn create_factory(id: FactoryID, xlsx_filepath: PathBuf) -> Box<dyn TemplateManager> {
    match id {
        FactoryID::ProfitAndLoss => Box::new(ProfitAndLossManager::new(xlsx_filepath)),
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    // CSVファイルとXLSXファイルのパスを取得する
    let csv_filepath = args.csv_filepath;
    let xlsx_filepath = args.xlsx_filepath;

    // ファクトリからTemplateManagerを生成して実行する
    let factory = create_factory(FactoryID::ProfitAndLoss, xlsx_filepath);
    factory.execute(csv_filepath)?;

    Ok(())
}
