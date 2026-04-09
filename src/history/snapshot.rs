use crate::model::OelDocument;

const DEFAULT_MAX_DEPTH: usize = 100;

pub struct UndoStack {
    undo: Vec<OelDocument>,
    redo: Vec<OelDocument>,
    max_depth: usize,
}

impl UndoStack {
    pub fn new() -> Self {
        Self { undo: Vec::new(), redo: Vec::new(), max_depth: DEFAULT_MAX_DEPTH }
    }

    /// Save the current document state so it can be restored by `undo`.
    /// Clears the redo stack.
    pub fn push_snapshot(&mut self, doc: &OelDocument) {
        if self.undo.len() >= self.max_depth {
            self.undo.remove(0);
        }
        self.undo.push(doc.clone());
        self.redo.clear();
    }

    /// Restore the previous state. `current` is pushed onto the redo stack.
    /// Returns `Some(previous_doc)` or `None` if nothing to undo.
    pub fn undo(&mut self, current: OelDocument) -> Option<OelDocument> {
        let prev = self.undo.pop()?;
        self.redo.push(current);
        Some(prev)
    }

    /// Restore the next state. `current` is pushed onto the undo stack.
    /// Returns `Some(next_doc)` or `None` if nothing to redo.
    pub fn redo(&mut self, current: OelDocument) -> Option<OelDocument> {
        let next = self.redo.pop()?;
        self.undo.push(current);
        Some(next)
    }

    pub fn has_undo(&self) -> bool {
        !self.undo.is_empty()
    }

    pub fn has_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    pub fn clear(&mut self) {
        self.undo.clear();
        self.redo.clear();
    }
}

impl Default for UndoStack {
    fn default() -> Self {
        Self::new()
    }
}
