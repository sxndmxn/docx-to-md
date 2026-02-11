use std::fmt::Write;

/// Render a grid of rows into a markdown table.
/// First row is used as the header.
pub fn render_table(rows: &[Vec<String>]) -> String {
    if rows.is_empty() {
        return String::new();
    }

    let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
    if col_count == 0 {
        return String::new();
    }

    let mut widths = vec![3usize; col_count];
    for row in rows {
        for (i, cell) in row.iter().enumerate() {
            let escaped = escape_table_cell(cell);
            if escaped.len() > widths[i] {
                widths[i] = escaped.len();
            }
        }
    }

    let mut out = String::new();

    // Header row
    out.push('|');
    if let Some(header) = rows.first() {
        for (i, w) in widths.iter().enumerate() {
            let cell = header.get(i).map_or("", String::as_str);
            let escaped = escape_table_cell(cell);
            let _ = write!(out, " {escaped:<w$} |", w = *w);
        }
    }
    out.push('\n');

    // Separator
    out.push('|');
    for w in &widths {
        let _ = write!(out, " {} |", "-".repeat(*w));
    }
    out.push('\n');

    // Data rows
    for row in rows.iter().skip(1) {
        out.push('|');
        for (i, w) in widths.iter().enumerate() {
            let cell = row.get(i).map_or("", String::as_str);
            let escaped = escape_table_cell(cell);
            let _ = write!(out, " {escaped:<w$} |", w = *w);
        }
        out.push('\n');
    }

    out
}

/// Escape pipe characters inside table cells
fn escape_table_cell(s: &str) -> String {
    s.replace('|', "\\|").replace('\n', " ").replace('\r', "")
}

/// Determine if a sheet is "tabular" (has consistent column usage)
/// vs "notes" (sparse cells, long text, single-column dumps).
///
/// Returns true if the data looks tabular.
pub fn is_tabular(rows: &[Vec<String>]) -> bool {
    if rows.len() < 2 {
        return false;
    }

    let col_count = rows.iter().map(Vec::len).max().unwrap_or(0);
    if col_count < 2 {
        return false;
    }

    // Check how many cells are non-empty
    let total_cells = rows.len() * col_count;
    let filled_cells: usize = rows
        .iter()
        .map(|row| row.iter().filter(|c| !c.trim().is_empty()).count())
        .sum();

    #[allow(clippy::cast_precision_loss)]
    let fill_ratio = filled_cells as f64 / total_cells as f64;

    // If less than 30% of cells are filled, it's notes-style
    fill_ratio > 0.3
}

/// Render sparse/notes-style sheet data as sequential paragraphs
pub fn render_notes(rows: &[Vec<String>]) -> String {
    let mut out = String::new();

    for row in rows {
        for cell in row {
            let trimmed = cell.trim();
            if !trimmed.is_empty() {
                out.push_str(trimmed);
                out.push_str("\n\n");
            }
        }
    }

    out
}
