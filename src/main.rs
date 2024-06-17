use clap::Parser;
use std::error::Error;
use std::path::PathBuf;
pub mod modules;
use modules::profit_and_loss_csv;

#[derive(Parser)]
struct Args {
    #[clap(name = "FILE")]
    csv_file: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    if let Some(csv_file) = args.csv_file {
        profit_and_loss_csv::execute(&csv_file);
    } else {
        return Err(
            "引数が不足しています。使用例: ./profit-and-loss-converter hogehoge.csv".into(),
        );
    }

    Ok(())
}
