use clap::Parser;
use modules::profit_and_loss;
use std::error::Error;
use std::path::PathBuf;
use umya_spreadsheet::reader::xlsx;

mod modules;

#[derive(Parser)]
struct Args {
    #[clap(name = "CSVFILE")]
    csv_filepath: Option<PathBuf>,

    #[clap(name = "XLSXFILE")]
    xlsx_filepath: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    match (args.csv_filepath, args.xlsx_filepath) {
        (Some(csv_filepath), Some(xlsx_filepath)) => {
            if let Err(err) = profit_and_loss::execute(&csv_filepath, &xlsx_filepath) {
                eprintln!("エラー: {}", err);
                std::process::exit(1);
            }
        }
        _ => {
            eprintln!("引数が不足しています。使用例: ./profit-and-loss-converter hogehoge.csv piyopiyo.xlsx");
            std::process::exit(1);
        }
    }

    Ok(())
}
