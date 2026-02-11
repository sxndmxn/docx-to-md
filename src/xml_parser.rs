use std::fmt::Write;
use std::path::Path;

use anyhow::Context;
use roxmltree::{Document, Node};

use crate::docx::DocxArchive;
use crate::embeds;
use crate::markdown;

const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_V: &str = "urn:schemas-microsoft-com:vml";

/// Convert the entire `document.xml` to markdown, resolving all embeds.
pub fn convert_document(docx: &DocxArchive, assets_dir: &Path) -> anyhow::Result<String> {
    let xml = docx.read_file_string("word/document.xml")?;
    let doc = Document::parse(&xml).context("Failed to parse document.xml")?;

    let mut output = String::new();
    let mut ctx = ConversionContext {
        docx,
        assets_dir,
        list_depth: 0,
    };

    let body = doc
        .descendants()
        .find(|n| n.has_tag_name((NS_W, "body")))
        .context("No w:body element found in document.xml")?;

    for child in body.children() {
        if child.has_tag_name((NS_W, "p")) {
            output.push_str(&convert_paragraph(&child, &mut ctx)?);
        } else if child.has_tag_name((NS_W, "tbl")) {
            output.push_str(&convert_table(&child, &mut ctx)?);
        }
    }

    Ok(output)
}

struct ConversionContext<'a> {
    docx: &'a DocxArchive,
    assets_dir: &'a Path,
    list_depth: usize,
}

fn convert_paragraph(node: &Node, ctx: &mut ConversionContext) -> anyhow::Result<String> {
    let mut line = String::new();

    let p_pr = node.children().find(|n| n.has_tag_name((NS_W, "pPr")));

    let heading_level = p_pr.as_ref().and_then(get_heading_level);
    let list_kind = p_pr.as_ref().and_then(get_list_kind);

    let mut text_parts = Vec::new();
    for child in node.children() {
        if child.has_tag_name((NS_W, "r")) {
            text_parts.push(convert_run(&child));
        } else if child.has_tag_name((NS_W, "hyperlink")) {
            text_parts.push(convert_hyperlink(&child, ctx));
        }
    }

    let text = text_parts.join("");
    let embed_content = extract_embeds_from_paragraph(node, ctx)?;

    if text.is_empty() && embed_content.is_empty() {
        return Ok("\n".to_string());
    }

    if let Some(level) = heading_level {
        let prefix = "#".repeat(level);
        line.push_str(&prefix);
        line.push(' ');
    }

    if let Some(kind) = &list_kind {
        let indent = "  ".repeat(ctx.list_depth);
        line.push_str(&indent);
        match kind {
            ListKind::Ordered => line.push_str("1. "),
            ListKind::Unordered => line.push_str("- "),
        }
    }

    line.push_str(&text);
    line.push('\n');

    if !embed_content.is_empty() {
        line.push('\n');
        line.push_str(&embed_content);
    }

    if heading_level.is_some() {
        line.push('\n');
    }

    Ok(line)
}

fn convert_run(node: &Node) -> String {
    let mut text = String::new();

    for child in node.children() {
        if child.has_tag_name((NS_W, "t")) {
            if let Some(t) = child.text() {
                text.push_str(t);
            }
        } else if child.has_tag_name((NS_W, "tab")) {
            text.push('\t');
        } else if child.has_tag_name((NS_W, "br")) {
            text.push('\n');
        }
    }

    if text.is_empty() {
        return String::new();
    }

    let r_pr = node.children().find(|n| n.has_tag_name((NS_W, "rPr")));

    if let Some(pr) = r_pr {
        let bold = pr.children().any(|n| n.has_tag_name((NS_W, "b")));
        let italic = pr.children().any(|n| n.has_tag_name((NS_W, "i")));
        let underline = pr
            .children()
            .any(|n| n.has_tag_name((NS_W, "u")) && n.attribute((NS_W, "val")) != Some("none"));
        let strike = pr.children().any(|n| n.has_tag_name((NS_W, "strike")));
        let code = pr.children().any(|n| {
            n.has_tag_name((NS_W, "rFonts"))
                && n.attribute((NS_W, "ascii")).is_some_and(is_monospace_font)
        });

        if code {
            text = format!("`{text}`");
        }
        if bold && italic {
            text = format!("***{text}***");
        } else if bold {
            text = format!("**{text}**");
        } else if italic {
            text = format!("*{text}*");
        }
        if underline {
            text = format!("<u>{text}</u>");
        }
        if strike {
            text = format!("~~{text}~~");
        }
    }

    text
}

fn convert_hyperlink(node: &Node, ctx: &ConversionContext) -> String {
    let mut text = String::new();

    for child in node.children() {
        if child.has_tag_name((NS_W, "r")) {
            text.push_str(&convert_run(&child));
        }
    }

    if let Some(rid) = node.attribute((NS_R, "id"))
        && let Some(rel) = ctx.docx.relationships.get(rid)
    {
        return format!("[{text}]({})", rel.target);
    }

    if let Some(anchor) = node.attribute((NS_W, "anchor")) {
        return format!("[{text}](#{anchor})");
    }

    text
}

fn convert_table(node: &Node, ctx: &mut ConversionContext) -> anyhow::Result<String> {
    let mut rows: Vec<Vec<String>> = Vec::new();

    for row_node in node.children() {
        if !row_node.has_tag_name((NS_W, "tr")) {
            continue;
        }

        let mut cells: Vec<String> = Vec::new();
        for cell_node in row_node.children() {
            if !cell_node.has_tag_name((NS_W, "tc")) {
                continue;
            }

            let mut cell_text = String::new();
            for p_node in cell_node.children() {
                if p_node.has_tag_name((NS_W, "p")) {
                    let para = convert_paragraph(&p_node, ctx)?;
                    cell_text.push_str(para.trim());
                    cell_text.push(' ');
                }
            }

            cells.push(cell_text.trim().to_string());
        }

        rows.push(cells);
    }

    if rows.is_empty() {
        return Ok(String::new());
    }

    let mut out = String::from("\n");
    out.push_str(&markdown::render_table(&rows));
    out.push('\n');

    Ok(out)
}

fn extract_embeds_from_paragraph(
    node: &Node,
    ctx: &mut ConversionContext,
) -> anyhow::Result<String> {
    let mut content = String::new();

    for desc in node.descendants() {
        if (desc.has_tag_name((NS_WP, "inline")) || desc.has_tag_name((NS_WP, "anchor")))
            && let Some(blip_content) = extract_blip(&desc, ctx)?
        {
            content.push_str(&blip_content);
        }

        if desc.has_tag_name((NS_W, "object")) {
            content.push_str(&extract_object(&desc, ctx)?);
        }

        // Handle VML imagedata only when NOT inside a w:object (those are handled above)
        if desc.has_tag_name((NS_V, "imagedata"))
            && !has_ancestor_tag(&desc, NS_W, "object")
            && let Some(rid) = desc.attribute((NS_R, "id"))
            && let Some(path) = ctx.docx.resolve_relationship(rid)
        {
            let path_str = path.to_string_lossy();
            if ctx.docx.has_file(&path_str) {
                let md_path = embeds::extract_image(ctx.docx, &path_str, ctx.assets_dir)?;
                let _ = write!(content, "![image]({md_path})\n\n");
            }
        }
    }

    Ok(content)
}

fn extract_blip(node: &Node, ctx: &mut ConversionContext) -> anyhow::Result<Option<String>> {
    let Some(blip) = node.descendants().find(|n| n.has_tag_name((NS_A, "blip"))) else {
        return Ok(None);
    };

    let Some(rid) = blip.attribute((NS_R, "embed")) else {
        return Ok(None);
    };

    let Some(path) = ctx.docx.resolve_relationship(rid) else {
        return Ok(None);
    };

    let path_str = path.to_string_lossy();

    if !ctx.docx.has_file(&path_str) {
        return Ok(None);
    }

    let alt = node
        .descendants()
        .find(|n| n.has_tag_name((NS_WP, "docPr")))
        .and_then(|n| n.attribute("descr"))
        .unwrap_or("image");

    let md_path = embeds::extract_image(ctx.docx, &path_str, ctx.assets_dir)?;
    Ok(Some(format!("![{alt}]({md_path})\n\n")))
}

fn extract_object(node: &Node, ctx: &mut ConversionContext) -> anyhow::Result<String> {
    for child in node.descendants() {
        let tag = child.tag_name().name();

        if (tag == "OLEObject" || tag == "oleObject")
            && let Some(rid) = child
                .attribute((NS_R, "id"))
                .or_else(|| child.attribute("r:id"))
            && let Some(path) = ctx.docx.resolve_relationship(rid)
        {
            let path_str = path.to_string_lossy();
            if ctx.docx.has_file(&path_str) {
                return handle_embedded_file(ctx, &path_str);
            }
        }
    }

    Ok(String::new())
}

fn handle_embedded_file(ctx: &ConversionContext, archive_path: &str) -> anyhow::Result<String> {
    let path = Path::new(archive_path);
    let ext = path.extension().map(|e| e.to_string_lossy().to_lowercase());

    match ext.as_deref() {
        Some("xlsx" | "xls") => embeds::extract_spreadsheet(ctx.docx, archive_path),
        Some("pdf") => embeds::extract_pdf(ctx.docx, archive_path),
        Some("csv") => {
            let text = ctx.docx.read_file_string(archive_path)?;
            embeds::csv_to_markdown(&text)
        }
        Some("bin") => embeds::extract_ole_object(ctx.docx, archive_path),
        Some("xml" | "xsd" | "xsl" | "xslt") => {
            let text = ctx.docx.read_file_string(archive_path)?;
            Ok(format!("```xml\n{text}\n```\n\n"))
        }
        Some(ext) if is_image_ext(ext) => {
            let md_path = embeds::extract_image(ctx.docx, archive_path, ctx.assets_dir)?;
            Ok(format!("![embedded]({md_path})\n\n"))
        }
        _ => Ok(format!(
            "> *Embedded file: {archive_path} (unsupported format)*\n\n"
        )),
    }
}

// --- Helpers ---

fn get_heading_level(p_pr: &Node) -> Option<usize> {
    let p_style = p_pr.children().find(|n| n.has_tag_name((NS_W, "pStyle")))?;

    let val = p_style.attribute((NS_W, "val"))?;

    if let Some(stripped) = val.strip_prefix("Heading") {
        return stripped.parse().ok();
    }

    let digits: String = val.chars().rev().take_while(char::is_ascii_digit).collect();
    let digits: String = digits.chars().rev().collect();

    if val.to_lowercase().contains("heading") || val.to_lowercase().contains("title") {
        let level: usize = digits.parse().unwrap_or(1);
        return Some(if level == 0 { 1 } else { level });
    }

    None
}

enum ListKind {
    Ordered,
    Unordered,
}

/// Detect list paragraphs via `w:numPr` (numbering XML) or paragraph style name.
fn get_list_kind(p_pr: &Node) -> Option<ListKind> {
    // Method 1: explicit w:numPr element (Word's native numbering)
    if let Some(num_pr) = p_pr.children().find(|n| n.has_tag_name((NS_W, "numPr"))) {
        let num_id = num_pr
            .children()
            .find(|n| n.has_tag_name((NS_W, "numId")))
            .and_then(|n| n.attribute((NS_W, "val")))
            .and_then(|v| v.parse::<i64>().ok());

        // numId 0 means "no list"
        if num_id == Some(0) {
            return None;
        }

        // Check ilvl for nesting depth hint; odd numId heuristic for ordered vs unordered
        return Some(if num_id.is_some_and(|id| id % 2 == 0) {
            ListKind::Ordered
        } else {
            ListKind::Unordered
        });
    }

    // Method 2: paragraph style name (e.g. ListBullet, ListNumber, List Paragraph)
    let style_val = p_pr
        .children()
        .find(|n| n.has_tag_name((NS_W, "pStyle")))
        .and_then(|n| n.attribute((NS_W, "val")))?;

    let lower = style_val.to_lowercase();

    if lower.contains("listnumber") || lower.contains("list number") {
        return Some(ListKind::Ordered);
    }

    if lower.contains("listbullet")
        || lower.contains("list bullet")
        || lower.contains("listparagraph")
        || lower.contains("list paragraph")
    {
        return Some(ListKind::Unordered);
    }

    None
}

fn is_monospace_font(font: &str) -> bool {
    let lower = font.to_lowercase();
    lower.contains("courier")
        || lower.contains("consolas")
        || lower.contains("mono")
        || lower.contains("menlo")
        || lower.contains("source code")
}

const IMAGE_EXTENSIONS: &[&str] = &[
    "png", "jpg", "jpeg", "gif", "bmp", "tiff", "tif", "emf", "wmf", "svg",
];

fn is_image_ext(ext: &str) -> bool {
    IMAGE_EXTENSIONS
        .iter()
        .any(|image_ext| ext.eq_ignore_ascii_case(image_ext))
}

/// Check if a node has an ancestor with the given namespace and tag name.
fn has_ancestor_tag(node: &Node, ns: &str, tag: &str) -> bool {
    let mut current = node.parent();
    while let Some(parent) = current {
        if parent.has_tag_name((ns, tag)) {
            return true;
        }
        current = parent.parent();
    }
    false
}
