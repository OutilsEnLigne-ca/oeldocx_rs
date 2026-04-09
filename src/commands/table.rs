use crate::model::{OelDocument, OelBlock, OelTable, OelTableRow, OelTableCell, OelParagraph, next_id};
use crate::cursor::{DocPosition, DocSelection};

/// Insert a new table at the current cursor block position.
pub fn insert_table(doc: &mut OelDocument, pos: &DocPosition, rows: usize, cols: usize) -> DocPosition {
    let table = OelTable::new(next_id(), rows, cols);
    let insert_idx = pos.block_idx + 1;
    doc.blocks.insert(insert_idx, OelBlock::Table(table));
    // Position cursor inside the first cell, first paragraph
    DocPosition::in_cell(insert_idx, 0, 0, 0, 0)
}

fn find_table_position(doc: &OelDocument, pos: &DocPosition) -> Option<(usize, usize, usize)> {
    // Returns (block_idx_of_table, row, col)
    match (pos.table_row, pos.table_col) {
        (Some(row), Some(col)) => Some((pos.block_idx, row, col)),
        _ => None,
    }
}

pub fn insert_row_above(doc: &mut OelDocument, pos: &DocPosition) {
    let Some((block_idx, row, _)) = find_table_position(doc, pos) else { return };
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        let cols = t.rows.first().map(|r| r.cells.len()).unwrap_or(1);
        t.rows.insert(row, OelTableRow::new(cols));
    }
}

pub fn insert_row_below(doc: &mut OelDocument, pos: &DocPosition) {
    let Some((block_idx, row, _)) = find_table_position(doc, pos) else { return };
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        let cols = t.rows.first().map(|r| r.cells.len()).unwrap_or(1);
        t.rows.insert(row + 1, OelTableRow::new(cols));
    }
}

pub fn insert_col_left(doc: &mut OelDocument, pos: &DocPosition) {
    let Some((block_idx, _, col)) = find_table_position(doc, pos) else { return };
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        for row in &mut t.rows {
            row.cells.insert(col, OelTableCell::new());
        }
    }
}

pub fn insert_col_right(doc: &mut OelDocument, pos: &DocPosition) {
    let Some((block_idx, _, col)) = find_table_position(doc, pos) else { return };
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        for row in &mut t.rows {
            row.cells.insert(col + 1, OelTableCell::new());
        }
    }
}

pub fn delete_row(doc: &mut OelDocument, pos: &DocPosition) -> Option<DocPosition> {
    let (block_idx, row, _) = find_table_position(doc, pos)?;
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        if t.rows.len() <= 1 {
            // Deleting the last row deletes the table
            doc.blocks.remove(block_idx);
            return Some(DocPosition::new(block_idx.saturating_sub(1), 0));
        }
        t.rows.remove(row);
        let new_row = row.min(t.rows.len() - 1);
        Some(DocPosition::in_cell(block_idx, new_row, 0, 0, 0))
    } else {
        None
    }
}

pub fn delete_col(doc: &mut OelDocument, pos: &DocPosition) -> Option<DocPosition> {
    let (block_idx, _, col) = find_table_position(doc, pos)?;
    if let Some(OelBlock::Table(t)) = doc.blocks.get_mut(block_idx) {
        let col_count = t.rows.first().map(|r| r.cells.len()).unwrap_or(0);
        if col_count <= 1 {
            // Deleting the last column deletes the table
            doc.blocks.remove(block_idx);
            return Some(DocPosition::new(block_idx.saturating_sub(1), 0));
        }
        for row in &mut t.rows {
            if col < row.cells.len() {
                row.cells.remove(col);
            }
        }
        let new_col = col.min(col_count - 2);
        Some(DocPosition::in_cell(block_idx, 0, new_col, 0, 0))
    } else {
        None
    }
}

pub fn delete_table(doc: &mut OelDocument, pos: &DocPosition) -> DocPosition {
    let block_idx = pos.block_idx;
    if block_idx < doc.blocks.len() {
        doc.blocks.remove(block_idx);
    }
    let new_idx = if doc.blocks.is_empty() {
        doc.blocks.push(OelBlock::Paragraph(OelParagraph::new(next_id())));
        0
    } else {
        block_idx.saturating_sub(1)
    };
    DocPosition::new(new_idx, 0)
}
