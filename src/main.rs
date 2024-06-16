use std::env;
use std::path::Path;

pub mod modules;
use modules::profit_and_loss_csv;

fn main() {
    let args: Vec<String> = env::args().collect();

    for (index, arg) in args.iter().enumerate() {
        println!("Argument {index}: {arg}");
        if index == 1 {
            let path = Path::new(arg);
            let result = profit_and_loss_csv::read_csv(path);

            println!("{result:?}");
        }
    }
}
