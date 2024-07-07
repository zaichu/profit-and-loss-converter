use std::env;
use std::fs;
use std::path::Path;
use std::path::PathBuf;

fn main() {
    // ビルド時に x86_64-w64-mingw32-gcc を使用する
    println!("cargo:rustc-env=CC=x86_64-w64-mingw32-gcc");
    println!("cargo:rustc-env=TARGET=x86_64-pc-windows-gnu");

    // プロジェクトのルートディレクトリのパスを取得
    let project_dir = env::var("CARGO_MANIFEST_DIR").unwrap();
    println!("Project directory: {project_dir}");

    // src/settings.json のパス
    let src_file = Path::new(&project_dir).join("src/settings.json");
    println!("Source file path: {:?}", src_file);

    // ビルド後のディレクトリのパスを取得
    let out_dir = env::var("OUT_DIR").unwrap();
    println!("Build output directory: {out_dir}");

    let mut out_dir = PathBuf::from(&out_dir);
    for _ in 0..3 {
        out_dir = out_dir.parent().unwrap().to_path_buf();
    }

    // コピー先のパス (ビルド後のディレクトリに settings.json をコピー)
    let dest_file = out_dir.join("settings.json");
    println!("Destination file path: {:?}", dest_file);

    // settings.json をビルド後のディレクトリにコピーする
    fs::copy(&src_file, &dest_file).expect("Failed to copy settings.json");
    println!("settings.json copied successfully!");

    let csv_dir = out_dir.join("csv");
    println!("Destination csv dirctory: {:?}", csv_dir);
    fs::create_dir(&csv_dir).expect("Failed to copy settings.json");
}
