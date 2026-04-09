use super::block::{OelBlock, OelParagraph, next_id};
use super::section::OelSectionProps;
use super::style::{OelParaProps, OelRunProps, OelStyle};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelDocument {
    pub blocks: Vec<OelBlock>,
    pub section: OelSectionProps,
    pub styles: HashMap<String, OelStyle>,
}

impl OelDocument {
    /// Empty document with a single blank paragraph.
    pub fn empty() -> Self {
        Self {
            blocks: vec![OelBlock::Paragraph(OelParagraph::new(next_id()))],
            section: OelSectionProps::default(),
            styles: default_styles(),
        }
    }

    pub fn word_count(&self) -> usize {
        self.iter_paragraphs()
            .map(|p| {
                let text = p.plain_text();
                text.split_whitespace().count()
            })
            .sum()
    }

    pub fn char_count(&self) -> usize {
        self.iter_paragraphs().map(|p| p.char_len()).sum()
    }

    /// Iterate all paragraphs in document order, including those inside table cells.
    pub fn iter_paragraphs(&self) -> impl Iterator<Item = &OelParagraph> {
        self.blocks.iter().flat_map(iter_block_paragraphs)
    }
}

fn iter_block_paragraphs(block: &OelBlock) -> Box<dyn Iterator<Item = &OelParagraph> + '_> {
    match block {
        OelBlock::Paragraph(p) => Box::new(std::iter::once(p)),
        OelBlock::Table(t) => Box::new(t.rows.iter().flat_map(|row| {
            row.cells
                .iter()
                .flat_map(|cell| cell.blocks.iter().flat_map(iter_block_paragraphs))
        })),
        OelBlock::PageBreak => Box::new(std::iter::empty()),
    }
}

fn default_styles() -> HashMap<String, OelStyle> {
    let mut styles = HashMap::new();

    styles.insert(
        "Normal".to_string(),
        OelStyle {
            id: "Normal".to_string(),
            name: "Normal".to_string(),
            run_props: OelRunProps {
                font_size: Some(22),
                font_family: Some("Roboto".to_string()),
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles.insert(
        "Title".to_string(),
        OelStyle {
            id: "Title".to_string(),
            name: "Title".to_string(),
            run_props: OelRunProps {
                font_size: Some(56),
                bold: true,
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles.insert(
        "Subtitle".to_string(),
        OelStyle {
            id: "Subtitle".to_string(),
            name: "Subtitle".to_string(),
            run_props: OelRunProps {
                font_size: Some(32),
                color: Some("666666".to_string()),
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles.insert(
        "Heading1".to_string(),
        OelStyle {
            id: "Heading1".to_string(),
            name: "Heading 1".to_string(),
            run_props: OelRunProps {
                font_size: Some(48),
                bold: true,
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles.insert(
        "Heading2".to_string(),
        OelStyle {
            id: "Heading2".to_string(),
            name: "Heading 2".to_string(),
            run_props: OelRunProps {
                font_size: Some(36),
                bold: true,
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles.insert(
        "Heading3".to_string(),
        OelStyle {
            id: "Heading3".to_string(),
            name: "Heading 3".to_string(),
            run_props: OelRunProps {
                font_size: Some(28),
                bold: true,
                ..Default::default()
            },
            para_props: OelParaProps::default(),
        },
    );

    styles
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{OelBlock, OelRun, OelTable, OelTableCell, OelTableRow};

    #[test]
    fn test_empty_document() {
        let doc = OelDocument::empty();
        assert_eq!(doc.blocks.len(), 1);
        assert!(matches!(doc.blocks[0], OelBlock::Paragraph(_)));
        assert!(doc.styles.contains_key("Normal"));
        assert_eq!(doc.word_count(), 0);
        assert_eq!(doc.char_count(), 0);
    }

    #[test]
    fn test_word_and_char_count() {
        let mut doc = OelDocument::empty();
        if let Some(OelBlock::Paragraph(p)) = doc.blocks.get_mut(0) {
            p.runs.push(OelRun::with_props(
                "Hello world ".to_string(),
                Default::default(),
            ));
        }

        assert_eq!(doc.word_count(), 2);
        assert_eq!(doc.char_count(), 12);
    }

    #[test]
    fn test_iter_paragraphs_with_table() {
        let mut doc = OelDocument::empty();

        let mut cell = OelTableCell {
            blocks: vec![],
            props: Default::default(),
        };
        let mut cell_para = OelParagraph::new("cell-p1".to_string());
        cell_para.runs.push(OelRun::with_props(
            "Table text".to_string(),
            Default::default(),
        ));
        cell.blocks.push(OelBlock::Paragraph(cell_para));

        let row = OelTableRow { cells: vec![cell] };
        let table = OelTable {
            id: "t1".to_string(),
            rows: vec![row],
            props: Default::default(),
        };

        doc.blocks.push(OelBlock::Table(table));

        let paras: Vec<&OelParagraph> = doc.iter_paragraphs().collect();
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].plain_text(), "");
        assert_eq!(paras[1].plain_text(), "Table text");

        assert_eq!(doc.word_count(), 2);
        assert_eq!(doc.char_count(), 10);
    }
}
