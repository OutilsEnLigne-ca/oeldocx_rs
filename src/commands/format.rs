use crate::model::{OelDocument, OelBlock, OelParagraph, OelRun, OelRunProps};
use crate::cursor::DocSelection;

/// Apply a run-property mutation to all runs within the selection.
/// For a collapsed selection (no range), toggles the property on the formatting
/// context (the controller tracks this as "pending format" — not implemented here,
/// handled by the controller).
pub fn apply_run_format(
    doc: &mut OelDocument,
    sel: &DocSelection,
    apply: &dyn Fn(&mut OelRunProps),
) {
    if sel.is_collapsed() {
        // Nothing to apply to a collapsed selection — caller handles pending format.
        return;
    }

    let (start, end) = sel.ordered();

    if start.block_idx == end.block_idx
        && start.table_row == end.table_row
        && start.table_col == end.table_col
        && start.inner_block_idx == end.inner_block_idx
    {
        // Single paragraph selection
        if let Some(para) = get_para_mut_raw(doc, start.block_idx, start.table_row, start.table_col, start.inner_block_idx) {
            apply_format_to_para(para, start.char_offset, end.char_offset, apply);
        }
    } else {
        // Multi-paragraph selection
        // First paragraph: from start.char_offset to end of paragraph
        if let Some(para) = get_para_mut_raw(doc, start.block_idx, start.table_row, start.table_col, start.inner_block_idx) {
            let len = para.char_len();
            apply_format_to_para(para, start.char_offset, len, apply);
        }

        // Middle paragraphs: entire paragraph
        let mut idx = start.block_idx + 1;
        while idx < end.block_idx {
            if let Some(OelBlock::Paragraph(para)) = doc.blocks.get_mut(idx) {
                let len = para.char_len();
                apply_format_to_para(para, 0, len, apply);
            }
            idx += 1;
        }

        // Last paragraph: from start to end.char_offset
        if let Some(para) = get_para_mut_raw(doc, end.block_idx, end.table_row, end.table_col, end.inner_block_idx) {
            apply_format_to_para(para, 0, end.char_offset, apply);
        }
    }

    // Normalize all affected paragraphs
    normalize_range(doc, start.block_idx, end.block_idx);
}

fn apply_format_to_para(
    para: &mut OelParagraph,
    char_start: usize,
    char_end: usize,
    apply: &dyn Fn(&mut OelRunProps),
) {
    let mut new_runs: Vec<OelRun> = Vec::new();
    let mut cursor = 0usize;

    for run in para.runs.drain(..) {
        let run_len = run.char_len();
        let run_start = cursor;
        let run_end = cursor + run_len;

        if run_end <= char_start || run_start >= char_end {
            // Entirely outside selection
            new_runs.push(run);
        } else {
            let sel_start_in_run = char_start.saturating_sub(run_start);
            let sel_end_in_run = (char_end - run_start).min(run_len);

            // Before part
            if sel_start_in_run > 0 {
                let text: String = run.text.chars().take(sel_start_in_run).collect();
                new_runs.push(OelRun::with_props(text, run.props.clone()));
            }

            // Selected part
            let text: String = run.text.chars()
                .skip(sel_start_in_run)
                .take(sel_end_in_run - sel_start_in_run)
                .collect();
            let mut props = run.props.clone();
            apply(&mut props);
            if !text.is_empty() {
                new_runs.push(OelRun::with_props(text, props));
            }

            // After part
            if sel_end_in_run < run_len {
                let text: String = run.text.chars().skip(sel_end_in_run).collect();
                new_runs.push(OelRun::with_props(text, run.props.clone()));
            }
        }

        cursor += run_len;
    }

    para.runs = new_runs;
}

fn normalize_range(doc: &mut OelDocument, from: usize, to: usize) {
    for idx in from..=to.min(doc.blocks.len().saturating_sub(1)) {
        if let Some(OelBlock::Paragraph(p)) = doc.blocks.get_mut(idx) {
            p.normalize_runs();
        }
    }
}

fn get_para_mut_raw<'a>(
    doc: &'a mut OelDocument,
    block_idx: usize,
    table_row: Option<usize>,
    table_col: Option<usize>,
    inner_block_idx: usize,
) -> Option<&'a mut OelParagraph> {
    let block = doc.blocks.get_mut(block_idx)?;
    match (table_row, table_col) {
        (None, _) => match block {
            OelBlock::Paragraph(p) => Some(p),
            _ => None,
        },
        (Some(row), Some(col)) => match block {
            OelBlock::Table(t) => {
                let cell = t.rows.get_mut(row)?.cells.get_mut(col)?;
                match cell.blocks.get_mut(inner_block_idx)? {
                    OelBlock::Paragraph(p) => Some(p),
                    _ => None,
                }
            }
            _ => None,
        },
        _ => None,
    }
}

// --- Convenience wrappers for individual format commands ---

pub fn set_bold(doc: &mut OelDocument, sel: &DocSelection, value: bool) {
    apply_run_format(doc, sel, &|p| p.bold = value);
}

pub fn set_italic(doc: &mut OelDocument, sel: &DocSelection, value: bool) {
    apply_run_format(doc, sel, &|p| p.italic = value);
}

pub fn set_underline(doc: &mut OelDocument, sel: &DocSelection, value: bool) {
    apply_run_format(doc, sel, &|p| p.underline = value);
}

pub fn set_strikethrough(doc: &mut OelDocument, sel: &DocSelection, value: bool) {
    apply_run_format(doc, sel, &|p| p.strikethrough = value);
}

pub fn set_font_size(doc: &mut OelDocument, sel: &DocSelection, half_points: u32) {
    apply_run_format(doc, sel, &move |p| p.font_size = Some(half_points));
}

pub fn set_font_family(doc: &mut OelDocument, sel: &DocSelection, family: &str) {
    let owned = family.to_string();
    apply_run_format(doc, sel, &move |p| p.font_family = Some(owned.clone()));
}

pub fn set_text_color(doc: &mut OelDocument, sel: &DocSelection, hex: &str) {
    let owned = hex.to_string();
    apply_run_format(doc, sel, &move |p| p.color = Some(owned.clone()));
}

pub fn set_highlight(doc: &mut OelDocument, sel: &DocSelection, hex: Option<&str>) {
    let owned = hex.map(|s| s.to_string());
    apply_run_format(doc, sel, &move |p| p.highlight = owned.clone());
}
