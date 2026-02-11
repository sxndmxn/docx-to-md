use std::collections::HashMap;
use std::io::Read;
use std::path::{Path, PathBuf};

use anyhow::Context;

/// A parsed relationship from `document.xml.rels`
#[derive(Debug, Clone)]
pub struct Relationship {
    pub target: String,
}

/// Represents an opened .docx archive with parsed relationships
pub struct DocxArchive {
    data: Vec<u8>,
    pub relationships: HashMap<String, Relationship>,
    file_list: Vec<String>,
}

impl DocxArchive {
    pub fn open(path: &Path) -> anyhow::Result<Self> {
        let data = std::fs::read(path)
            .with_context(|| format!("Failed to read docx file: {}", path.display()))?;

        let file_list = {
            let cursor = std::io::Cursor::new(&data);
            let archive =
                zip::ZipArchive::new(cursor).context("Failed to open docx as ZIP archive")?;
            (0..archive.len())
                .filter_map(|i| archive.name_for_index(i).map(String::from))
                .collect::<Vec<_>>()
        };

        let relationships = Self::parse_relationships(&data)?;

        Ok(Self {
            data,
            relationships,
            file_list,
        })
    }

    /// Read a file from the archive by its path
    pub fn read_file(&self, name: &str) -> anyhow::Result<Vec<u8>> {
        let cursor = std::io::Cursor::new(&self.data);
        let mut archive = zip::ZipArchive::new(cursor).context("Failed to open ZIP archive")?;
        let mut file = archive
            .by_name(name)
            .with_context(|| format!("File not found in archive: {name}"))?;
        let mut buf = Vec::new();
        file.read_to_end(&mut buf)?;
        Ok(buf)
    }

    /// Read a file as UTF-8 string
    pub fn read_file_string(&self, name: &str) -> anyhow::Result<String> {
        let bytes = self.read_file(name)?;
        String::from_utf8(bytes).with_context(|| format!("File is not valid UTF-8: {name}"))
    }

    /// Check if a file exists in the archive
    pub fn has_file(&self, name: &str) -> bool {
        self.file_list.iter().any(|f| f == name)
    }

    /// Get the target path for a relationship ID, resolved relative to `word/`
    pub fn resolve_relationship(&self, rel_id: &str) -> Option<PathBuf> {
        self.relationships.get(rel_id).map(|rel| {
            rel.target
                .strip_prefix('/')
                .map_or_else(|| PathBuf::from("word").join(&rel.target), PathBuf::from)
        })
    }

    fn parse_relationships(data: &[u8]) -> anyhow::Result<HashMap<String, Relationship>> {
        let cursor = std::io::Cursor::new(data);
        let mut archive = zip::ZipArchive::new(cursor).context("Failed to open ZIP archive")?;

        let rels_path = "word/_rels/document.xml.rels";
        let mut rels_map = HashMap::new();

        let Ok(mut file) = archive.by_name(rels_path) else {
            return Ok(rels_map);
        };

        let mut xml = String::new();
        file.read_to_string(&mut xml)?;

        let doc = roxmltree::Document::parse(&xml).context("Failed to parse document.xml.rels")?;

        for node in doc.descendants() {
            if node.tag_name().name() == "Relationship"
                && let (Some(id), Some(target)) = (node.attribute("Id"), node.attribute("Target"))
            {
                rels_map.insert(
                    id.to_string(),
                    Relationship {
                        target: target.to_string(),
                    },
                );
            }
        }

        Ok(rels_map)
    }
}
