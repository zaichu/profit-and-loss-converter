use clap::Parser;
use modules::csv_reader::CSVReader;
use modules::excel_writer::ExcelWriter;
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
    let profit_and_loss = CSVReader::read_profit_and_loss(&csv_filepath)?;
    ExcelWriter::update_sheet(profit_and_loss, &xlsx_filepath)?;
    Ok(())
}
