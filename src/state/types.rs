use serde::{Deserialize, Serialize};
use crate::model::{Alignment, ListType};
use crate::render::RenderDocument;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CursorState {
    pub block_idx: usize,
    pub table_row: Option<usize>,
    pub table_col: Option<usize>,
    pub inner_block_idx: usize,
    pub char_offset: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SelectionState {
    pub anchor: CursorState,
    pub focus: CursorState,
    pub is_collapsed: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FormatState {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size: u32,
    pub font_family: String,
    pub color: String,
    pub highlight: Option<String>,
    pub alignment: Alignment,
    pub list_type: Option<ListType>,
    pub is_in_table: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct HistoryState {
    pub has_undo: bool,
    pub has_redo: bool,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct DocumentInfo {
    pub word_count: usize,
    pub char_count: usize,
    pub filename: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EditorState {
    pub document: RenderDocument,
    pub cursor: CursorState,
    pub selection: SelectionState,
    pub format: FormatState,
    pub history: HistoryState,
    pub document_info: DocumentInfo,
}
