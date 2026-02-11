use std::collections::HashMap;
use std::fmt::Write;
use std::io::{Cursor, Read};
use std::path::Path;

use anyhow::Context;
use calamine::{Data, Reader, Xlsx, open_workbook_from_rs};

use crate::docx::DocxArchive;
use crate::markdown;

/// Extract an image from the archive and copy it to the assets directory.
/// Returns the relative markdown path like `assets/image1.png`.
pub fn extract_image(
    docx: &DocxArchive,
    archive_path: &str,
    assets_dir: &Path,
) -> anyhow::Result<String> {
    let data = docx.read_file(archive_path)?;

    let filename = Path::new(archive_path).file_name().map_or_else(
        || "image.bin".to_string(),
        |n| n.to_string_lossy().into_owned(),
    );

    let dest = assets_dir.join(&filename);
    std::fs::write(dest, &data)?;

    Ok(format!("assets/{filename}"))
}

/// Extract and convert an embedded spreadsheet to markdown.
pub fn extract_spreadsheet(docx: &DocxArchive, archive_path: &str) -> anyhow::Result<String> {
    let data = docx.read_file(archive_path)?;
    spreadsheet_to_markdown(&data, archive_path)
}

/// Extract an OLE-embedded object — try to find an xlsx inside the OLE container.
pub fn extract_ole_object(docx: &DocxArchive, archive_path: &str) -> anyhow::Result<String> {
    let data = docx.read_file(archive_path)?;
    let cursor = Cursor::new(&data);

    let mut ole = cfb::CompoundFile::open(cursor)
        .with_context(|| format!("Failed to open OLE object: {archive_path}"))?;

    let stream_candidates = [
        "/Package",
        "/\x01Ole10Native",
        "/Workbook",
        "/Book",
        "/Contents",
    ];

    for stream_name in &stream_candidates {
        if let Ok(mut stream) = ole.open_stream(stream_name) {
            let mut buf = Vec::new();
            stream.read_to_end(&mut buf)?;

            if let Ok(result) = spreadsheet_to_markdown(&buf, archive_path) {
                return Ok(result);
            }

            if let Ok(text) = String::from_utf8(buf)
                && text.contains(',')
                && let Ok(result) = csv_to_markdown(&text)
            {
                return Ok(result);
            }
        }
    }

    Ok(format!(
        "> *Embedded OLE object: {archive_path} (could not extract content)*\n\n"
    ))
}

/// Extract text from an embedded PDF.
pub fn extract_pdf(docx: &DocxArchive, archive_path: &str) -> anyhow::Result<String> {
    let data = docx.read_file(archive_path)?;

    let text = pdf_extract::extract_text_from_mem(&data)
        .with_context(|| format!("Failed to extract text from PDF: {archive_path}"))?;

    let trimmed = text.trim();
    if trimmed.is_empty() {
        return Ok(format!(
            "> *Embedded PDF: {archive_path} (no extractable text)*\n\n"
        ));
    }

    let filename = Path::new(archive_path).file_name().map_or_else(
        || "document.pdf".to_string(),
        |n| n.to_string_lossy().into_owned(),
    );

    Ok(format!(
        "---\n\n**Embedded PDF: {filename}**\n\n{trimmed}\n\n---\n\n"
    ))
}

/// Convert CSV text to markdown.
pub fn csv_to_markdown(text: &str) -> anyhow::Result<String> {
    let mut reader = csv::ReaderBuilder::new()
        .has_headers(false)
        .from_reader(text.as_bytes());

    let mut rows: Vec<Vec<String>> = Vec::new();
    for result in reader.records() {
        let record = result?;
        rows.push(record.iter().map(String::from).collect());
    }

    if rows.is_empty() {
        return Ok(String::new());
    }

    Ok(markdown::render_table(&rows))
}

/// Convert xlsx bytes to markdown.
fn spreadsheet_to_markdown(data: &[u8], source_name: &str) -> anyhow::Result<String> {
    let cursor = Cursor::new(data);
    let mut workbook: Xlsx<_> = open_workbook_from_rs(cursor)
        .with_context(|| format!("Failed to open spreadsheet: {source_name}"))?;

    #[allow(clippy::implicit_clone)]
    let sheet_names = workbook.sheet_names().to_vec();
    let mut output = String::new();

    let comments = extract_xlsx_comments(data);
    let hyperlinks = extract_xlsx_hyperlinks(data);
    let merged = extract_xlsx_merged_cells(data);

    for sheet_name in &sheet_names {
        let Ok(range) = workbook.worksheet_range(sheet_name) else {
            continue;
        };

        if range.is_empty() {
            continue;
        }

        if sheet_names.len() > 1 {
            let _ = write!(output, "**Sheet: {sheet_name}**\n\n");
        }

        let mut rows: Vec<Vec<String>> = Vec::new();
        let merged_for_sheet = merged.get(sheet_name.as_str());

        for (row_idx, row) in range.rows().enumerate() {
            let mut cells: Vec<String> = Vec::new();
            for (col_idx, cell) in row.iter().enumerate() {
                if is_merged_secondary(merged_for_sheet, row_idx, col_idx) {
                    cells.push(String::new());
                    continue;
                }

                let mut value = cell_to_string(cell);

                if let Some(url) = hyperlinks
                    .get(sheet_name.as_str())
                    .and_then(|m| m.get(&(row_idx, col_idx)))
                {
                    if value.is_empty() {
                        value.clone_from(url);
                    } else {
                        value = format!("[{value}]({url})");
                    }
                }

                if let Some(comment) = comments
                    .get(sheet_name.as_str())
                    .and_then(|m| m.get(&(row_idx, col_idx)))
                {
                    if value.is_empty() {
                        value = format!("*(Note: {comment})*");
                    } else {
                        value = format!("{value} *(Note: {comment})*");
                    }
                }

                cells.push(value);
            }
            rows.push(cells);
        }

        while rows.last().is_some_and(|r| r.iter().all(String::is_empty)) {
            rows.pop();
        }

        if rows.is_empty() {
            continue;
        }

        if markdown::is_tabular(&rows) {
            output.push_str(&markdown::render_table(&rows));
        } else {
            output.push_str(&markdown::render_notes(&rows));
        }

        output.push('\n');
    }

    Ok(output)
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.clone(),
        Data::Float(f) => {
            if (*f - f.round()).abs() < f64::EPSILON {
                #[allow(clippy::cast_possible_truncation)]
                let i = *f as i64;
                format!("{i}")
            } else {
                format!("{f}")
            }
        }
        Data::Int(i) => format!("{i}"),
        Data::Bool(b) => format!("{b}"),
        Data::DateTime(dt) => format!("{dt}"),
        Data::Error(e) => format!("{e:?}"),
    }
}

// --- xlsx internal XML parsing for comments, hyperlinks, merged cells ---

struct MergeRange {
    start_row: usize,
    start_col: usize,
    end_row: usize,
    end_col: usize,
}

/// Try to extract comments from xlsx internal XML.
/// On any failure, returns an empty map (comments are best-effort).
fn extract_xlsx_comments(data: &[u8]) -> HashMap<String, HashMap<(usize, usize), String>> {
    let mut result = HashMap::new();

    let cursor = Cursor::new(data);
    let Ok(mut archive) = zip::ZipArchive::new(cursor) else {
        return result;
    };

    let comment_files: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let name = archive.name_for_index(i)?.to_string();
            if name.starts_with("xl/comments") && has_xml_ext(&name) {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for comment_file in &comment_files {
        let Ok(xml) = read_zip_string(&mut archive, comment_file) else {
            continue;
        };

        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };

        let sheet_key = comment_file
            .trim_start_matches("xl/")
            .trim_end_matches(".xml")
            .to_string();

        let mut comments = HashMap::new();
        for node in doc.descendants() {
            if node.tag_name().name() == "comment"
                && let Some(cell_ref) = node.attribute("ref")
            {
                let text: String = node
                    .descendants()
                    .filter(|n| n.tag_name().name() == "t")
                    .filter_map(|n| n.text())
                    .collect::<Vec<_>>()
                    .join(" ");

                if let Some(pos) = parse_cell_ref(cell_ref) {
                    comments.insert(pos, text.trim().to_string());
                }
            }
        }

        if !comments.is_empty() {
            result.insert(sheet_key, comments);
        }
    }

    result
}

/// Try to extract hyperlinks from xlsx sheet relationship files.
fn extract_xlsx_hyperlinks(data: &[u8]) -> HashMap<String, HashMap<(usize, usize), String>> {
    let mut result = HashMap::new();

    let cursor = Cursor::new(data);
    let Ok(mut archive) = zip::ZipArchive::new(cursor) else {
        return result;
    };

    let sheet_files: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let name = archive.name_for_index(i)?.to_string();
            if name.starts_with("xl/worksheets/sheet") && has_xml_ext(&name) {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for sheet_file in &sheet_files {
        let Ok(xml) = read_zip_string(&mut archive, sheet_file) else {
            continue;
        };

        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };

        let rels_path = sheet_file.replace("xl/worksheets/", "xl/worksheets/_rels/") + ".rels";

        let rels_map = parse_rels_file(&mut archive, &rels_path);

        let sheet_key = sheet_file
            .trim_start_matches("xl/worksheets/")
            .trim_end_matches(".xml")
            .to_string();

        let mut links = HashMap::new();

        for node in doc.descendants() {
            if node.tag_name().name() == "hyperlink" {
                let cell_ref = node.attribute("ref");
                let rid = node.attribute((
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    "id",
                ));

                let url = rid
                    .and_then(|id| rels_map.get(id).cloned())
                    .or_else(|| node.attribute("location").map(String::from));

                if let (Some(cell_ref), Some(url)) = (cell_ref, url)
                    && let Some(pos) = parse_cell_ref(cell_ref)
                {
                    links.insert(pos, url);
                }
            }
        }

        if !links.is_empty() {
            result.insert(sheet_key, links);
        }
    }

    result
}

/// Try to extract merged cell ranges from xlsx sheet XML.
fn extract_xlsx_merged_cells(data: &[u8]) -> HashMap<String, Vec<MergeRange>> {
    let mut result = HashMap::new();

    let cursor = Cursor::new(data);
    let Ok(mut archive) = zip::ZipArchive::new(cursor) else {
        return result;
    };

    let sheet_files: Vec<String> = (0..archive.len())
        .filter_map(|i| {
            let name = archive.name_for_index(i)?.to_string();
            if name.starts_with("xl/worksheets/sheet") && has_xml_ext(&name) {
                Some(name)
            } else {
                None
            }
        })
        .collect();

    for sheet_file in &sheet_files {
        let Ok(xml) = read_zip_string(&mut archive, sheet_file) else {
            continue;
        };

        let Ok(doc) = roxmltree::Document::parse(&xml) else {
            continue;
        };

        let sheet_key = sheet_file
            .trim_start_matches("xl/worksheets/")
            .trim_end_matches(".xml")
            .to_string();

        let mut ranges = Vec::new();

        for node in doc.descendants() {
            if node.tag_name().name() == "mergeCell"
                && let Some(range_ref) = node.attribute("ref")
                && let Some(range) = parse_merge_ref(range_ref)
            {
                ranges.push(range);
            }
        }

        if !ranges.is_empty() {
            result.insert(sheet_key, ranges);
        }
    }

    result
}

fn read_zip_string(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    name: &str,
) -> anyhow::Result<String> {
    let mut file = archive.by_name(name)?;
    let mut s = String::new();
    file.read_to_string(&mut s)?;
    Ok(s)
}

fn parse_rels_file(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
    path: &str,
) -> HashMap<String, String> {
    let mut map = HashMap::new();

    let Ok(xml) = read_zip_string(archive, path) else {
        return map;
    };

    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return map;
    };

    for node in doc.descendants() {
        if node.tag_name().name() == "Relationship"
            && let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target"))
        {
            map.insert(id.to_string(), target.to_string());
        }
    }

    map
}

/// Parse a cell reference like "A1" into (row, col) zero-indexed.
fn parse_cell_ref(cell_ref: &str) -> Option<(usize, usize)> {
    let mut col_str = String::new();
    let mut row_str = String::new();

    for ch in cell_ref.chars() {
        if ch.is_ascii_alphabetic() {
            col_str.push(ch.to_ascii_uppercase());
        } else if ch.is_ascii_digit() {
            row_str.push(ch);
        }
    }

    let row: usize = row_str.parse::<usize>().ok()?.checked_sub(1)?;
    let col = col_letters_to_index(&col_str)?;

    Some((row, col))
}

/// Convert column letters (A, B, ..., Z, AA, AB, ...) to zero-indexed column number.
fn col_letters_to_index(letters: &str) -> Option<usize> {
    if letters.is_empty() {
        return None;
    }
    let mut idx: usize = 0;
    for ch in letters.chars() {
        let digit = (ch as usize).checked_sub('A' as usize)?;
        idx = idx.checked_mul(26)?.checked_add(digit + 1)?;
    }
    Some(idx - 1)
}

/// Parse a merge reference like "A1:C3" into a `MergeRange`.
fn parse_merge_ref(range_ref: &str) -> Option<MergeRange> {
    let (start, end) = range_ref.split_once(':')?;
    let (start_row, start_col) = parse_cell_ref(start)?;
    let (end_row, end_col) = parse_cell_ref(end)?;
    Some(MergeRange {
        start_row,
        start_col,
        end_row,
        end_col,
    })
}

/// Check if a cell at (row, col) is inside a merge range but is NOT the top-left origin.
fn is_merged_secondary(merged: Option<&Vec<MergeRange>>, row: usize, col: usize) -> bool {
    let Some(ranges) = merged else {
        return false;
    };
    ranges.iter().any(|r| {
        row >= r.start_row
            && row <= r.end_row
            && col >= r.start_col
            && col <= r.end_col
            && (row != r.start_row || col != r.start_col)
    })
}

fn has_xml_ext(name: &str) -> bool {
    Path::new(name)
        .extension()
        .is_some_and(|ext| ext.eq_ignore_ascii_case("xml"))
}
