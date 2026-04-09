use serde::{Deserialize, Serialize};
use super::block::{OelBlock, OelParagraph, next_id};
use super::section::OelSectionProps;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelDocument {
    pub blocks: Vec<OelBlock>,
    pub section: OelSectionProps,
}

impl OelDocument {
    /// Empty document with a single blank paragraph.
    pub fn empty() -> Self {
        Self {
            blocks: vec![OelBlock::Paragraph(OelParagraph::new(next_id()))],
            section: OelSectionProps::default(),
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
        self.iter_paragraphs()
            .map(|p| p.char_len())
            .sum()
    }

    /// Iterate all paragraphs in document order, including those inside table cells.
    pub fn iter_paragraphs(&self) -> impl Iterator<Item = &OelParagraph> {
        self.blocks.iter().flat_map(iter_block_paragraphs)
    }
}

fn iter_block_paragraphs(block: &OelBlock) -> Box<dyn Iterator<Item = &OelParagraph> + '_> {
    match block {
        OelBlock::Paragraph(p) => Box::new(std::iter::once(p)),
        OelBlock::Table(t) => Box::new(
            t.rows.iter().flat_map(|row| {
                row.cells.iter().flat_map(|cell| {
                    cell.blocks.iter().flat_map(iter_block_paragraphs)
                })
            }),
        ),
        OelBlock::PageBreak => Box::new(std::iter::empty()),
    }
}
