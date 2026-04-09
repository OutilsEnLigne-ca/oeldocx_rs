use crate::model::{OelDocument, OelBlock, OelParagraph, next_id};
use crate::cursor::DocPosition;

pub fn insert_page_break(doc: &mut OelDocument, pos: &DocPosition) -> DocPosition {
    let insert_idx = pos.block_idx + 1;
    doc.blocks.insert(insert_idx, OelBlock::PageBreak);
    // Place cursor in a new paragraph after the page break
    let para_idx = insert_idx + 1;
    doc.blocks.insert(para_idx, OelBlock::Paragraph(OelParagraph::new(next_id())));
    DocPosition::new(para_idx, 0)
}

/// `width` and `height` in twips.
pub fn set_page_size(doc: &mut OelDocument, width: u32, height: u32) {
    doc.section.page_width = width;
    doc.section.page_height = height;
}

/// All margins in twips.
pub fn set_margins(doc: &mut OelDocument, top: u32, right: u32, bottom: u32, left: u32) {
    doc.section.margin_top = top;
    doc.section.margin_right = right;
    doc.section.margin_bottom = bottom;
    doc.section.margin_left = left;
}
