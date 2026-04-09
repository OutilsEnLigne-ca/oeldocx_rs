use serde::{Deserialize, Serialize};

/// An address within the document pointing to a character position.
///
/// For top-level paragraphs: `table_row = None, table_col = None, inner_block_idx = 0`.
/// For table cells: `table_row = Some(r), table_col = Some(c)`, `inner_block_idx` selects
/// which block inside the cell, `char_offset` is within that paragraph.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct DocPosition {
    pub block_idx: usize,
    pub table_row: Option<usize>,
    pub table_col: Option<usize>,
    /// Block index within a table cell (ignored for top-level paragraphs).
    pub inner_block_idx: usize,
    /// Character offset within the paragraph (Unicode scalar values, not bytes).
    pub char_offset: usize,
}

impl DocPosition {
    pub fn new(block_idx: usize, char_offset: usize) -> Self {
        Self {
            block_idx,
            table_row: None,
            table_col: None,
            inner_block_idx: 0,
            char_offset,
        }
    }

    pub fn in_cell(block_idx: usize, row: usize, col: usize, inner_block_idx: usize, char_offset: usize) -> Self {
        Self {
            block_idx,
            table_row: Some(row),
            table_col: Some(col),
            inner_block_idx,
            char_offset,
        }
    }

    pub fn is_in_table(&self) -> bool {
        self.table_row.is_some()
    }

    /// Returns `(block_idx, table_row, table_col, inner_block_idx, char_offset)` as flat tuple.
    pub fn as_tuple(&self) -> (usize, Option<usize>, Option<usize>, usize, usize) {
        (self.block_idx, self.table_row, self.table_col, self.inner_block_idx, self.char_offset)
    }
}

impl Default for DocPosition {
    fn default() -> Self {
        Self::new(0, 0)
    }
}
