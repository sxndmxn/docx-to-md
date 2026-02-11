#![deny(
    clippy::all,
    clippy::pedantic,
    clippy::unwrap_used,
    clippy::expect_used
)]
#![deny(unsafe_code)]
#![warn(clippy::nursery)]

mod docx;
mod embeds;
mod markdown;
mod xml_parser;

use std::path::PathBuf;

use clap::Parser;

#[derive(Parser)]
#[command(name = "docx-to-md", about = "Convert DOCX files to Markdown")]
struct Cli {
    /// Path to the .docx file
    input: PathBuf,

    /// Output directory
    #[arg(short, long, default_value = "output")]
    output: PathBuf,
}

fn main() -> anyhow::Result<()> {
    let cli = Cli::parse();

    if !cli.input.exists() {
        anyhow::bail!("Input file does not exist: {}", cli.input.display());
    }

    let assets_dir = cli.output.join("assets");
    std::fs::create_dir_all(&assets_dir)?;

    let docx = docx::DocxArchive::open(&cli.input)?;
    let md = xml_parser::convert_document(&docx, &assets_dir)?;

    std::fs::write(cli.output.join("document.md"), md)?;

    eprintln!(
        "Done. Output written to {}",
        cli.output.join("document.md").display()
    );

    Ok(())
}
