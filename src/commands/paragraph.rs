use crate::model::{OelDocument, OelBlock, OelParagraph, OelParaProps, Alignment, ListType};
use crate::cursor::DocSelection;

/// Apply a paragraph-property mutation to all paragraphs touched by the selection.
pub fn apply_para_format(
    doc: &mut OelDocument,
    sel: &DocSelection,
    apply: &dyn Fn(&mut OelParaProps),
) {
    let (start, end) = sel.ordered();
    let from = start.block_idx;
    let to = end.block_idx;

    for idx in from..=to.min(doc.blocks.len().saturating_sub(1)) {
        match doc.blocks.get_mut(idx) {
            Some(OelBlock::Paragraph(p)) => apply(&mut p.props),
            Some(OelBlock::Table(t)) => {
                // Apply to all paragraphs in the table if the whole table is selected
                if idx > from && idx < to {
                    for row in &mut t.rows {
                        for cell in &mut row.cells {
                            for block in &mut cell.blocks {
                                if let OelBlock::Paragraph(p) = block {
                                    apply(&mut p.props);
                                }
                            }
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

pub fn set_alignment(doc: &mut OelDocument, sel: &DocSelection, alignment: Alignment) {
    apply_para_format(doc, sel, &move |p| p.alignment = Some(alignment.clone()));
}

pub fn set_indent(doc: &mut OelDocument, sel: &DocSelection, level: u32) {
    apply_para_format(doc, sel, &move |p| p.indent_level = level);
}

pub fn increase_indent(doc: &mut OelDocument, sel: &DocSelection) {
    apply_para_format(doc, sel, &|p| p.indent_level = p.indent_level.saturating_add(1));
}

pub fn decrease_indent(doc: &mut OelDocument, sel: &DocSelection) {
    apply_para_format(doc, sel, &|p| p.indent_level = p.indent_level.saturating_sub(1));
}

/// numId used for user-created bullet lists (matches BULLET_NUM_ID in to_docx.rs)
const BULLET_NUM_ID: u32 = 1;
/// numId used for user-created numbered lists (matches NUMBERED_NUM_ID in to_docx.rs)
const NUMBERED_NUM_ID: u32 = 2;

pub fn toggle_bullet_list(doc: &mut OelDocument, sel: &DocSelection) {
    let (start, end) = sel.ordered();
    let all_bullets = (start.block_idx..=end.block_idx).all(|idx| {
        matches!(
            doc.blocks.get(idx),
            Some(OelBlock::Paragraph(p)) if p.props.list_type == Some(ListType::Bullet)
        )
    });

    if all_bullets {
        apply_para_format(doc, sel, &|p| { p.list_type = None; p.num_id = None; });
    } else {
        apply_para_format(doc, sel, &|p| { p.list_type = Some(ListType::Bullet); p.num_id = Some(BULLET_NUM_ID); });
    }
}

pub fn toggle_numbered_list(doc: &mut OelDocument, sel: &DocSelection) {
    let (start, end) = sel.ordered();
    let all_numbered = (start.block_idx..=end.block_idx).all(|idx| {
        matches!(
            doc.blocks.get(idx),
            Some(OelBlock::Paragraph(p)) if p.props.list_type == Some(ListType::Numbered)
        )
    });

    if all_numbered {
        apply_para_format(doc, sel, &|p| { p.list_type = None; p.num_id = None; });
    } else {
        apply_para_format(doc, sel, &|p| { p.list_type = Some(ListType::Numbered); p.num_id = Some(NUMBERED_NUM_ID); });
    }
}

pub fn set_line_spacing(doc: &mut OelDocument, sel: &DocSelection, multiplier: f32) {
    apply_para_format(doc, sel, &move |p| p.line_spacing = Some(multiplier));
}

pub fn set_paragraph_spacing(doc: &mut OelDocument, sel: &DocSelection, before: u32, after: u32) {
    apply_para_format(doc, sel, &move |p| {
        p.spacing_before = Some(before);
        p.spacing_after = Some(after);
    });
}

pub fn apply_style(doc: &mut OelDocument, sel: &DocSelection, style_id: &str) {
    let owned = style_id.to_string();
    apply_para_format(doc, sel, &move |p| p.style_id = Some(owned.clone()));
}
