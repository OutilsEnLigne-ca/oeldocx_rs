use super::position::DocPosition;

#[derive(Debug, Clone)]
pub struct DocSelection {
    /// Where the selection started (fixed end during shift+arrow extension).
    pub anchor: DocPosition,
    /// Where the cursor currently is.
    pub focus: DocPosition,
}

impl DocSelection {
    pub fn collapsed(pos: DocPosition) -> Self {
        Self { anchor: pos.clone(), focus: pos }
    }

    pub fn new(anchor: DocPosition, focus: DocPosition) -> Self {
        Self { anchor, focus }
    }

    pub fn is_collapsed(&self) -> bool {
        self.anchor == self.focus
    }

    /// Move to a new collapsed position.
    pub fn set_collapsed(&mut self, pos: DocPosition) {
        self.anchor = pos.clone();
        self.focus = pos;
    }

    /// Returns (start, end) in document order (by block_idx then char_offset).
    /// For same-block selections, anchor/focus are ordered by char_offset.
    pub fn ordered(&self) -> (&DocPosition, &DocPosition) {
        let a = &self.anchor;
        let f = &self.focus;
        if a.block_idx < f.block_idx {
            return (a, f);
        }
        if a.block_idx > f.block_idx {
            return (f, a);
        }
        // Same block — order by inner block, then char_offset
        if a.inner_block_idx < f.inner_block_idx {
            return (a, f);
        }
        if a.inner_block_idx > f.inner_block_idx {
            return (f, a);
        }
        if a.char_offset <= f.char_offset { (a, f) } else { (f, a) }
    }
}

impl Default for DocSelection {
    fn default() -> Self {
        Self::collapsed(DocPosition::default())
    }
}
