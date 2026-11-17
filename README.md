# docx-to-md

`docx-to-md` converts Word `.docx` files to Markdown.
It extracts document text, tables, images, hyperlinks, and supported embedded files.

## Install

Install Rust and Cargo.
Run this command from the repository root:

```bash
cargo install --path . --locked
```

## Use

```bash
docx-to-md report.docx --output converted-report
docx-to-md --help
```

The tool writes `document.md` and an `assets/` directory inside the output directory.
The default output directory is `output`.
Choose a separate output directory for each document to prevent output files from replacing earlier files.

## Supported embedded files

The converter handles supported Excel workbooks, PDF documents, and XML or XSD content.
The extracted Markdown represents document content, rather than the exact Word page layout.

## Development

```bash
cargo check --locked
```
