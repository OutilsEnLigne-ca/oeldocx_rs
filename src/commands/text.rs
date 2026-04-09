use crate::model::{OelDocument, OelBlock, OelParagraph, OelRun, OelRunProps, next_id};
use crate::cursor::DocPosition;

/// Get a mutable reference to the paragraph addressed by `pos`.
pub fn get_para_mut<'a>(doc: &'a mut OelDocument, pos: &DocPosition) -> Option<&'a mut OelParagraph> {
    let block = doc.blocks.get_mut(pos.block_idx)?;
    match (pos.table_row, pos.table_col) {
        (None, _) => match block {
            OelBlock::Paragraph(p) => Some(p),
            _ => None,
        },
        (Some(row), Some(col)) => match block {
            OelBlock::Table(t) => {
                let cell = t.rows.get_mut(row)?.cells.get_mut(col)?;
                match cell.blocks.get_mut(pos.inner_block_idx)? {
                    OelBlock::Paragraph(p) => Some(p),
                    _ => None,
                }
            }
            _ => None,
        },
        _ => None,
    }
}

/// Get an immutable reference to the paragraph addressed by `pos`.
pub fn get_para<'a>(doc: &'a OelDocument, pos: &DocPosition) -> Option<&'a OelParagraph> {
    let block = doc.blocks.get(pos.block_idx)?;
    match (pos.table_row, pos.table_col) {
        (None, _) => match block {
            OelBlock::Paragraph(p) => Some(p),
            _ => None,
        },
        (Some(row), Some(col)) => match block {
            OelBlock::Table(t) => {
                let cell = t.rows.get(row)?.cells.get(col)?;
                match cell.blocks.get(pos.inner_block_idx)? {
                    OelBlock::Paragraph(p) => Some(p),
                    _ => None,
                }
            }
            _ => None,
        },
        _ => None,
    }
}

/// Resolve (run_idx, offset_in_run) for a char_offset within a paragraph.
fn resolve_run_offset(para: &OelParagraph, char_offset: usize) -> (usize, usize) {
    let mut accumulated = 0usize;
    for (i, run) in para.runs.iter().enumerate() {
        let len = run.char_len();
        if accumulated + len >= char_offset || i + 1 == para.runs.len() {
            return (i, char_offset.saturating_sub(accumulated).min(len));
        }
        accumulated += len;
    }
    (0, 0)
}

/// Insert `text` at the cursor position.
pub fn insert_text(doc: &mut OelDocument, pos: &mut DocPosition, text: &str) {
    let Some(para) = get_para_mut(doc, pos) else { return };
    insert_text_into_para(para, pos.char_offset, text);
    pos.char_offset += text.chars().count();
}

pub fn insert_text_into_para(para: &mut OelParagraph, char_offset: usize, text: &str) {
    if para.runs.is_empty() {
        para.runs.push(OelRun::new(text));
        return;
    }

    let (run_idx, offset_in_run) = resolve_run_offset(para, char_offset);
    let run = &para.runs[run_idx];
    let before: String = run.text.chars().take(offset_in_run).collect();
    let after: String = run.text.chars().skip(offset_in_run).collect();
    let props = run.props.clone();

    para.runs[run_idx] = OelRun::with_props(format!("{}{}{}", before, text, after), props);
}

/// Delete one character before the cursor.
pub fn delete_backward(doc: &mut OelDocument, pos: &mut DocPosition) {
    if pos.char_offset == 0 {
        if pos.table_row.is_none() && pos.block_idx > 0 {
            merge_with_previous(doc, pos);
        }
        return;
    }

    let Some(para) = get_para_mut(doc, pos) else { return };
    delete_char_in_para(para, pos.char_offset - 1);
    pos.char_offset -= 1;
}

/// Delete one character after the cursor.
pub fn delete_forward(doc: &mut OelDocument, pos: &mut DocPosition) {
    let para_len = get_para(doc, pos).map(|p| p.char_len()).unwrap_or(0);
    if pos.char_offset >= para_len {
        if pos.table_row.is_none() {
            merge_with_next(doc, pos);
        }
        return;
    }

    let Some(para) = get_para_mut(doc, pos) else { return };
    delete_char_in_para(para, pos.char_offset);
}

fn delete_char_in_para(para: &mut OelParagraph, char_offset: usize) {
    let (run_idx, offset_in_run) = resolve_run_offset(para, char_offset);
    let chars: Vec<char> = para.runs[run_idx].text.chars().collect();
    if offset_in_run < chars.len() {
        let mut new_chars = chars;
        new_chars.remove(offset_in_run);
        para.runs[run_idx].text = new_chars.into_iter().collect();
        if para.runs[run_idx].text.is_empty() {
            para.runs.remove(run_idx);
        }
    }
}

/// Split the paragraph at `char_offset` — before stays, after goes to a new paragraph.
pub fn insert_newline(doc: &mut OelDocument, pos: &mut DocPosition) {
    if pos.table_row.is_some() {
        insert_newline_in_cell(doc, pos);
        return;
    }

    let block_idx = pos.block_idx;
    let char_offset = pos.char_offset;

    let (before_runs, after_runs, inherited_props) = {
        let Some(OelBlock::Paragraph(para)) = doc.blocks.get(block_idx) else { return };
        split_runs_at(para, char_offset)
    };

    let new_para = OelParagraph {
        id: next_id(),
        props: inherited_props,
        runs: after_runs,
    };

    if let Some(OelBlock::Paragraph(para)) = doc.blocks.get_mut(block_idx) {
        para.runs = before_runs;
        para.normalize_runs();
    }

    doc.blocks.insert(block_idx + 1, OelBlock::Paragraph(new_para));
    pos.block_idx += 1;
    pos.char_offset = 0;
}

fn insert_newline_in_cell(doc: &mut OelDocument, pos: &mut DocPosition) {
    // Use inner fn returning Option<()> so we can use `?` cleanly.
    fn inner(doc: &mut OelDocument, pos: &mut DocPosition) -> Option<()> {
        let row = pos.table_row?;
        let col = pos.table_col?;
        let inner_idx = pos.inner_block_idx;
        let char_offset = pos.char_offset;

        let (before_runs, after_runs, inherited_props) = {
            let table = match doc.blocks.get(pos.block_idx)? {
                OelBlock::Table(t) => t,
                _ => return None,
            };
            let para = match table.rows.get(row)?.cells.get(col)?.blocks.get(inner_idx)? {
                OelBlock::Paragraph(p) => p,
                _ => return None,
            };
            split_runs_at(para, char_offset)
        };

        let new_para = OelBlock::Paragraph(OelParagraph {
            id: next_id(),
            props: inherited_props,
            runs: after_runs,
        });

        let table = match doc.blocks.get_mut(pos.block_idx)? {
            OelBlock::Table(t) => t,
            _ => return None,
        };
        let cell = table.rows.get_mut(row)?.cells.get_mut(col)?;

        if let Some(OelBlock::Paragraph(p)) = cell.blocks.get_mut(inner_idx) {
            p.runs = before_runs;
            p.normalize_runs();
        }

        cell.blocks.insert(inner_idx + 1, new_para);
        pos.inner_block_idx += 1;
        pos.char_offset = 0;
        Some(())
    }

    inner(doc, pos);
}

/// Split runs of a paragraph at `char_offset`.
/// Returns (before_runs, after_runs, cloned_para_props).
fn split_runs_at(para: &OelParagraph, char_offset: usize) -> (Vec<OelRun>, Vec<OelRun>, crate::model::OelParaProps) {
    let mut before = Vec::new();
    let mut after = Vec::new();
    let mut accumulated = 0usize;

    for run in &para.runs {
        let len = run.char_len();
        let run_start = accumulated;
        let run_end = accumulated + len;

        if run_end <= char_offset {
            before.push(run.clone());
        } else if run_start >= char_offset {
            after.push(run.clone());
        } else {
            let split_at = char_offset - run_start;
            let text_before: String = run.text.chars().take(split_at).collect();
            let text_after: String = run.text.chars().skip(split_at).collect();
            if !text_before.is_empty() {
                before.push(OelRun::with_props(text_before, run.props.clone()));
            }
            if !text_after.is_empty() {
                after.push(OelRun::with_props(text_after, run.props.clone()));
            }
        }

        accumulated += len;
    }

    (before, after, para.props.clone())
}

fn merge_with_previous(doc: &mut OelDocument, pos: &mut DocPosition) {
    let curr_idx = pos.block_idx;
    let prev_idx = curr_idx - 1;

    let (prev_len, runs_to_append) = {
        let prev = match doc.blocks.get(prev_idx) {
            Some(OelBlock::Paragraph(p)) => p,
            _ => return,
        };
        let curr = match doc.blocks.get(curr_idx) {
            Some(OelBlock::Paragraph(p)) => p,
            _ => return,
        };
        (prev.char_len(), curr.runs.clone())
    };

    if let Some(OelBlock::Paragraph(prev)) = doc.blocks.get_mut(prev_idx) {
        prev.runs.extend(runs_to_append);
        prev.normalize_runs();
    }

    doc.blocks.remove(curr_idx);
    pos.block_idx = prev_idx;
    pos.char_offset = prev_len;
}

fn merge_with_next(doc: &mut OelDocument, pos: &mut DocPosition) {
    let curr_idx = pos.block_idx;
    let next_idx = curr_idx + 1;
    if next_idx >= doc.blocks.len() {
        return;
    }

    let runs_to_append = match doc.blocks.get(next_idx) {
        Some(OelBlock::Paragraph(p)) => p.runs.clone(),
        _ => return,
    };

    if let Some(OelBlock::Paragraph(curr)) = doc.blocks.get_mut(curr_idx) {
        curr.runs.extend(runs_to_append);
        curr.normalize_runs();
    }

    doc.blocks.remove(next_idx);
}
