use once_cell::sync::Lazy;
use serde::{Deserialize, Serialize};
use std::error::Error;
use std::fs::File;
use std::io::BufReader;

#[derive(Debug, Serialize, Deserialize)]
pub struct Settings {
    pub formats: std::collections::HashMap<String, String>,
    pub colors: std::collections::HashMap<String, String>,
    pub headers: std::collections::HashMap<String, String>,
    pub sheet_title: String,
    pub tax_rate: f64,
    pub start_row: u32,
    pub start_col: u32,
}

impl Settings {
    pub fn load_from_file(path: &str) -> Result<Self, Box<dyn Error>> {
        let file = File::open(path)?;
        let reader = BufReader::new(file);
        let settings: Settings = serde_json::from_reader(reader)?;
        Ok(settings)
    }
}

pub static SETTINGS: Lazy<Settings> =
    Lazy::new(|| Settings::load_from_file("settings.json").expect("Failed to load settings"));
