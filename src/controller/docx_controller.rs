use crate::commands::{
    format as fmt_cmd, paragraph as para_cmd, section as sec_cmd, table as tbl_cmd, text as txt_cmd,
};
use crate::convert::{SerializeError, docx_to_oel, oel_to_bytes, oel_to_render};
use crate::cursor::{DocPosition, DocSelection};
use crate::history::UndoStack;
use crate::model::{Alignment, OelBlock, OelDocument, OelParagraph, OelRunProps, next_id};
use crate::render::{DEFAULT_COLOR, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, HALFPT_TO_PT};
use crate::state::{
    CursorState, DocumentInfo, EditorState, FormatState, HistoryState, SelectionState,
};

#[derive(Debug)]
pub enum ControllerError {
    ParseError(String),
    SerializeError(String),
    NoDocument,
}

impl std::fmt::Display for ControllerError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            ControllerError::ParseError(s) => write!(f, "parse error: {s}"),
            ControllerError::SerializeError(s) => write!(f, "serialize error: {s}"),
            ControllerError::NoDocument => write!(f, "no document loaded"),
        }
    }
}

/// The main stateful editing engine for a DOCX document.
///
/// `DocxController` encapsulates the mutable document model, the current user
/// cursor/selection, and the history (undo/redo) stack. It receives editing
/// commands, mutates the document, and produces state snapshots for the frontend.
pub struct DocxController {
    pub document: OelDocument,
    pub selection: DocSelection,
    pub undo_stack: UndoStack,
    pub filename: Option<String>,
    /// Formatting to apply to the next inserted character (when selection is collapsed).
    pub pending_format: Option<OelRunProps>,
    /// Whether the previous command was a text-input (for undo batching).
    last_was_text_input: bool,
}

impl DocxController {
    pub fn new() -> Self {
        Self {
            document: OelDocument::empty(),
            selection: DocSelection::default(),
            undo_stack: UndoStack::new(),
            filename: None,
            pending_format: None,
            last_was_text_input: false,
        }
    }

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    pub fn load(
        &mut self,
        bytes: &[u8],
        filename: Option<String>,
    ) -> Result<EditorState, ControllerError> {
        let docx =
            docx_rs::read_docx(bytes).map_err(|e| ControllerError::ParseError(e.to_string()))?;

        self.document = docx_to_oel(&docx);
        self.selection = DocSelection::default();
        self.undo_stack.clear();
        self.filename = filename;
        self.pending_format = None;
        self.last_was_text_input = false;

        Ok(self.build_state())
    }

    pub fn new_document(&mut self) -> EditorState {
        self.document = OelDocument::empty();
        self.selection = DocSelection::default();
        self.undo_stack.clear();
        self.filename = None;
        self.pending_format = None;
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn serialize(&self) -> Result<Vec<u8>, ControllerError> {
        oel_to_bytes(&self.document).map_err(|e| ControllerError::SerializeError(e.to_string()))
    }

    pub fn get_state(&self) -> EditorState {
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Cursor / Selection
    // -------------------------------------------------------------------------

    pub fn set_selection(
        &mut self,
        anchor_block: usize,
        anchor_row: Option<usize>,
        anchor_col: Option<usize>,
        anchor_inner: usize,
        anchor_offset: usize,
        focus_block: usize,
        focus_row: Option<usize>,
        focus_col: Option<usize>,
        focus_inner: usize,
        focus_offset: usize,
    ) -> EditorState {
        self.selection = DocSelection::new(
            DocPosition {
                block_idx: anchor_block,
                table_row: anchor_row,
                table_col: anchor_col,
                inner_block_idx: anchor_inner,
                char_offset: anchor_offset,
            },
            DocPosition {
                block_idx: focus_block,
                table_row: focus_row,
                table_col: focus_col,
                inner_block_idx: focus_inner,
                char_offset: focus_offset,
            },
        );
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Text input
    // -------------------------------------------------------------------------

    pub fn insert_text(&mut self, text: &str) -> EditorState {
        self.snapshot_for_text();
        let mut pos = self.selection.focus.clone();
        txt_cmd::insert_text(&mut self.document, &mut pos, text);
        self.selection.set_collapsed(pos);
        self.last_was_text_input = true;
        self.build_state()
    }

    pub fn delete_backward(&mut self) -> EditorState {
        self.snapshot_for_text();
        let mut pos = self.selection.focus.clone();
        txt_cmd::delete_backward(&mut self.document, &mut pos);
        self.selection.set_collapsed(pos);
        self.last_was_text_input = true;
        self.build_state()
    }

    pub fn delete_forward(&mut self) -> EditorState {
        self.snapshot_for_text();
        let mut pos = self.selection.focus.clone();
        txt_cmd::delete_forward(&mut self.document, &mut pos);
        self.selection.set_collapsed(pos);
        self.last_was_text_input = true;
        self.build_state()
    }

    pub fn insert_newline(&mut self) -> EditorState {
        self.snapshot();
        let mut pos = self.selection.focus.clone();
        txt_cmd::insert_newline(&mut self.document, &mut pos);
        self.selection.set_collapsed(pos);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Character formatting
    // -------------------------------------------------------------------------

    pub fn set_bold(&mut self, value: bool) -> EditorState {
        self.snapshot();
        fmt_cmd::set_bold(&mut self.document, &self.selection, value);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_italic(&mut self, value: bool) -> EditorState {
        self.snapshot();
        fmt_cmd::set_italic(&mut self.document, &self.selection, value);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_underline(&mut self, value: bool) -> EditorState {
        self.snapshot();
        fmt_cmd::set_underline(&mut self.document, &self.selection, value);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_strikethrough(&mut self, value: bool) -> EditorState {
        self.snapshot();
        fmt_cmd::set_strikethrough(&mut self.document, &self.selection, value);
        self.last_was_text_input = false;
        self.build_state()
    }

    /// `half_points`: font size in OOXML half-points (24 = 12pt).
    pub fn set_font_size(&mut self, half_points: u32) -> EditorState {
        self.snapshot();
        fmt_cmd::set_font_size(&mut self.document, &self.selection, half_points);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_font_family(&mut self, family: &str) -> EditorState {
        self.snapshot();
        fmt_cmd::set_font_family(&mut self.document, &self.selection, family);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_text_color(&mut self, hex: &str) -> EditorState {
        self.snapshot();
        fmt_cmd::set_text_color(&mut self.document, &self.selection, hex);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_highlight(&mut self, hex: Option<&str>) -> EditorState {
        self.snapshot();
        fmt_cmd::set_highlight(&mut self.document, &self.selection, hex);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Paragraph formatting
    // -------------------------------------------------------------------------

    pub fn set_alignment(&mut self, alignment: Alignment) -> EditorState {
        self.snapshot();
        para_cmd::set_alignment(&mut self.document, &self.selection, alignment);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_indent(&mut self, level: u32) -> EditorState {
        self.snapshot();
        para_cmd::set_indent(&mut self.document, &self.selection, level);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn increase_indent(&mut self) -> EditorState {
        self.snapshot();
        para_cmd::increase_indent(&mut self.document, &self.selection);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn decrease_indent(&mut self) -> EditorState {
        self.snapshot();
        para_cmd::decrease_indent(&mut self.document, &self.selection);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn toggle_bullet_list(&mut self) -> EditorState {
        self.snapshot();
        para_cmd::toggle_bullet_list(&mut self.document, &self.selection);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn toggle_numbered_list(&mut self) -> EditorState {
        self.snapshot();
        para_cmd::toggle_numbered_list(&mut self.document, &self.selection);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_line_spacing(&mut self, multiplier: f32) -> EditorState {
        self.snapshot();
        para_cmd::set_line_spacing(&mut self.document, &self.selection, multiplier);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_paragraph_spacing(&mut self, before: u32, after: u32) -> EditorState {
        self.snapshot();
        para_cmd::set_paragraph_spacing(&mut self.document, &self.selection, before, after);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn apply_style(&mut self, style_id: &str) -> EditorState {
        self.snapshot();
        para_cmd::apply_style(&mut self.document, &self.selection, style_id);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Table operations
    // -------------------------------------------------------------------------

    pub fn insert_table(&mut self, rows: u32, cols: u32) -> EditorState {
        self.snapshot();
        let new_pos = tbl_cmd::insert_table(
            &mut self.document,
            &self.selection.focus,
            rows as usize,
            cols as usize,
        );
        self.selection.set_collapsed(new_pos);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn insert_row_above(&mut self) -> EditorState {
        self.snapshot();
        tbl_cmd::insert_row_above(&mut self.document, &self.selection.focus);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn insert_row_below(&mut self) -> EditorState {
        self.snapshot();
        tbl_cmd::insert_row_below(&mut self.document, &self.selection.focus);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn insert_col_left(&mut self) -> EditorState {
        self.snapshot();
        tbl_cmd::insert_col_left(&mut self.document, &self.selection.focus);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn insert_col_right(&mut self) -> EditorState {
        self.snapshot();
        tbl_cmd::insert_col_right(&mut self.document, &self.selection.focus);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn delete_row(&mut self) -> EditorState {
        self.snapshot();
        if let Some(new_pos) = tbl_cmd::delete_row(&mut self.document, &self.selection.focus) {
            self.selection.set_collapsed(new_pos);
        }
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn delete_col(&mut self) -> EditorState {
        self.snapshot();
        if let Some(new_pos) = tbl_cmd::delete_col(&mut self.document, &self.selection.focus) {
            self.selection.set_collapsed(new_pos);
        }
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn delete_table(&mut self) -> EditorState {
        self.snapshot();
        let new_pos = tbl_cmd::delete_table(&mut self.document, &self.selection.focus);
        self.selection.set_collapsed(new_pos);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Section / page
    // -------------------------------------------------------------------------

    pub fn insert_page_break(&mut self) -> EditorState {
        self.snapshot();
        let new_pos = sec_cmd::insert_page_break(&mut self.document, &self.selection.focus);
        self.selection.set_collapsed(new_pos);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_page_size(&mut self, width_twips: u32, height_twips: u32) -> EditorState {
        self.snapshot();
        sec_cmd::set_page_size(&mut self.document, width_twips, height_twips);
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn set_margins(&mut self, top: u32, right: u32, bottom: u32, left: u32) -> EditorState {
        self.snapshot();
        sec_cmd::set_margins(&mut self.document, top, right, bottom, left);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // History
    // -------------------------------------------------------------------------

    pub fn undo(&mut self) -> EditorState {
        if let Some(prev) = self.undo_stack.undo(self.document.clone()) {
            self.document = prev;
            // Reset cursor to document start — could be smarter in the future
            self.selection = DocSelection::default();
        }
        self.last_was_text_input = false;
        self.build_state()
    }

    pub fn redo(&mut self) -> EditorState {
        if let Some(next) = self.undo_stack.redo(self.document.clone()) {
            self.document = next;
            self.selection = DocSelection::default();
        }
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Clipboard (plain text only in Phase 1)
    // -------------------------------------------------------------------------

    pub fn copy(&self) -> String {
        if self.selection.is_collapsed() {
            return String::new();
        }
        let (start, end) = self.selection.ordered();
        if start.block_idx == end.block_idx {
            if let Some(OelBlock::Paragraph(para)) = self.document.blocks.get(start.block_idx) {
                let full = para.plain_text();
                let chars: Vec<char> = full.chars().collect();
                chars[start.char_offset..end.char_offset.min(chars.len())]
                    .iter()
                    .collect()
            } else {
                String::new()
            }
        } else {
            // Multi-block: collect text from each paragraph, joined by newlines
            let mut parts = Vec::new();
            for idx in start.block_idx..=end.block_idx {
                if let Some(OelBlock::Paragraph(para)) = self.document.blocks.get(idx) {
                    let full = para.plain_text();
                    let chars: Vec<char> = full.chars().collect();
                    let from = if idx == start.block_idx {
                        start.char_offset
                    } else {
                        0
                    };
                    let to = if idx == end.block_idx {
                        end.char_offset
                    } else {
                        chars.len()
                    };
                    parts.push(chars[from..to.min(chars.len())].iter().collect::<String>());
                }
            }
            parts.join("\n")
        }
    }

    pub fn paste(&mut self, text: &str) -> EditorState {
        self.snapshot();
        let mut pos = self.selection.focus.clone();
        txt_cmd::insert_text(&mut self.document, &mut pos, text);
        self.selection.set_collapsed(pos);
        self.last_was_text_input = false;
        self.build_state()
    }

    // -------------------------------------------------------------------------
    // Internal helpers
    // -------------------------------------------------------------------------

    fn snapshot(&mut self) {
        self.undo_stack.push_snapshot(&self.document);
        self.last_was_text_input = false;
    }

    /// Push snapshot only when transitioning from non-text to text input.
    fn snapshot_for_text(&mut self) {
        if !self.last_was_text_input {
            self.undo_stack.push_snapshot(&self.document);
        }
    }

    fn build_state(&self) -> EditorState {
        let render_doc = oel_to_render(&self.document);
        let cursor = self.cursor_state();
        let selection = self.selection_state();
        let format = self.format_at_cursor();
        let history = HistoryState {
            has_undo: self.undo_stack.has_undo(),
            has_redo: self.undo_stack.has_redo(),
        };
        let document_info = DocumentInfo {
            word_count: self.document.word_count(),
            char_count: self.document.char_count(),
            filename: self.filename.clone(),
        };
        EditorState {
            document: render_doc,
            cursor,
            selection,
            format,
            history,
            document_info,
        }
    }

    fn cursor_state(&self) -> CursorState {
        let f = &self.selection.focus;
        CursorState {
            block_idx: f.block_idx,
            table_row: f.table_row,
            table_col: f.table_col,
            inner_block_idx: f.inner_block_idx,
            char_offset: f.char_offset,
        }
    }

    fn selection_state(&self) -> SelectionState {
        let to_cursor = |p: &DocPosition| CursorState {
            block_idx: p.block_idx,
            table_row: p.table_row,
            table_col: p.table_col,
            inner_block_idx: p.inner_block_idx,
            char_offset: p.char_offset,
        };
        SelectionState {
            anchor: to_cursor(&self.selection.anchor),
            focus: to_cursor(&self.selection.focus),
            is_collapsed: self.selection.is_collapsed(),
        }
    }

    fn format_at_cursor(&self) -> FormatState {
        let pos = &self.selection.focus;
        let is_in_table = pos.is_in_table();

        // Resolve the run props at the cursor
        let run_props = self.run_props_at(pos);

        let para_props = self.para_props_at(pos);

        let style_run_props = para_props
            .and_then(|pp| pp.style_id.as_ref())
            .and_then(|id| self.document.styles.get(id))
            .map(|s| &s.run_props);

        let (alignment, list_type) = match para_props {
            Some(pp) => (pp.alignment.clone(), pp.list_type.clone()),
            None => (crate::model::Alignment::Left, None),
        };

        let bold = run_props.bold || style_run_props.map_or(false, |sp| sp.bold);
        let italic = run_props.italic || style_run_props.map_or(false, |sp| sp.italic);
        let underline = run_props.underline || style_run_props.map_or(false, |sp| sp.underline);
        let strikethrough =
            run_props.strikethrough || style_run_props.map_or(false, |sp| sp.strikethrough);

        let font_size = run_props
            .font_size
            .or_else(|| style_run_props.and_then(|sp| sp.font_size));
        let font_family = run_props
            .font_family
            .clone()
            .or_else(|| style_run_props.and_then(|sp| sp.font_family.clone()));
        let color = run_props
            .color
            .clone()
            .or_else(|| style_run_props.and_then(|sp| sp.color.clone()));

        FormatState {
            bold,
            italic,
            underline,
            strikethrough,
            font_size: font_size
                .map(|s| s / 2) // half-points → points for state
                .unwrap_or(11),
            font_family: font_family.unwrap_or_else(|| DEFAULT_FONT_FAMILY.to_string()),
            color: color.unwrap_or_else(|| DEFAULT_COLOR.to_string()),
            highlight: run_props.highlight,
            alignment,
            list_type,
            is_in_table,
            current_style_id: para_props.and_then(|pp| pp.style_id.clone()),
        }
    }

    fn run_props_at(&self, pos: &DocPosition) -> OelRunProps {
        let block = self.document.blocks.get(pos.block_idx);
        let para = match (pos.table_row, pos.table_col, block) {
            (None, _, Some(OelBlock::Paragraph(p))) => Some(p),
            (Some(row), Some(col), Some(OelBlock::Table(t))) => t
                .rows
                .get(row)
                .and_then(|r| r.cells.get(col))
                .and_then(|c| c.blocks.get(pos.inner_block_idx))
                .and_then(|b| b.as_paragraph()),
            _ => None,
        };

        let Some(para) = para else {
            return OelRunProps::default();
        };

        // Find the run at or just before the cursor
        let mut accumulated = 0usize;
        let mut best: Option<&crate::model::OelRunProps> = None;
        for run in &para.runs {
            let len = run.char_len();
            if accumulated + len >= pos.char_offset {
                best = Some(&run.props);
                break;
            }
            accumulated += len;
            best = Some(&run.props);
        }
        best.cloned().unwrap_or_default()
    }

    fn para_props_at(&self, pos: &DocPosition) -> Option<&crate::model::OelParaProps> {
        let block = self.document.blocks.get(pos.block_idx)?;
        match (pos.table_row, pos.table_col) {
            (None, _) => match block {
                OelBlock::Paragraph(p) => Some(&p.props),
                _ => None,
            },
            (Some(row), Some(col)) => match block {
                OelBlock::Table(t) => t
                    .rows
                    .get(row)?
                    .cells
                    .get(col)?
                    .blocks
                    .get(pos.inner_block_idx)?
                    .as_paragraph()
                    .map(|p| &p.props),
                _ => None,
            },
            _ => None,
        }
    }
}

impl Default for DocxController {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_format_at_cursor_with_style() {
        let mut controller = DocxController::new();

        controller.apply_style("Heading1");

        let state = controller.get_state();
        assert_eq!(state.format.current_style_id.as_deref(), Some("Heading1"));
        assert_eq!(state.format.bold, true);
        assert_eq!(state.format.font_size, 24);
    }
}
