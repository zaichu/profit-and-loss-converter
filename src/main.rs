use clap::Parser;
use modules::profit_and_loss::lib::ProfitAndLossManager;
use modules::template_pattern::TemplateManager;
use std::error::Error;
use std::path::PathBuf;

mod modules;

#[derive(Parser)]
struct Args {
    #[clap(name = "CSVFILE")]
    csv_filepath: Option<PathBuf>,
    #[clap(name = "XLSXFILE")]
    xlsx_filepath: Option<PathBuf>,
}

enum FactoryID {
    ProfitAndLoss,
}

fn create_factory(
    id: FactoryID,
    csv_filepath: PathBuf,
    xlsx_filepath: PathBuf,
) -> Box<dyn TemplateManager> {
    match id {
        FactoryID::ProfitAndLoss => {
            Box::new(ProfitAndLossManager::new(csv_filepath, xlsx_filepath))
        }
    }
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    if let (Some(csv_filepath), Some(xlsx_filepath)) = (args.csv_filepath, args.xlsx_filepath) {
        execute(csv_filepath, xlsx_filepath)?;
    } else {
        eprintln!(
            "引数が不足しています。使用例: ./profit-and-loss-converter hogehoge.csv piyopiyo.xlsx"
        );
        std::process::exit(1);
    }

    Ok(())
}

fn execute(csv_filepath: PathBuf, xlsx_filepath: PathBuf) -> Result<(), Box<dyn Error>> {
    let factory = create_factory(FactoryID::ProfitAndLoss, csv_filepath, xlsx_filepath);
    factory.excute()?;
    Ok(())
}
