use clap::Parser;
use modules::profit_and_loss;
use std::error::Error;
use std::path::PathBuf;

mod modules;

#[derive(Parser)]
struct Args {
    #[clap(name = "FILE")]
    csv_file: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    match args.csv_file {
        Some(csv_file) => {
            if let Err(err) = profit_and_loss::execute(&csv_file) {
                eprintln!("エラー: {}", err);
                std::process::exit(1);
            }
        }
        None => {
            eprintln!("引数が不足しています。使用例: ./profit-and-loss-converter hogehoge.csv");
            std::process::exit(1);
        }
    }

    Ok(())
}
