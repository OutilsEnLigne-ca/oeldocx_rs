use super::style::{OelParaProps, OelRunProps, OelTableCellProps, OelTableProps};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum OelWrappingMode {
    Inline,
    Square,
    Tight,
    Through,
    TopAndBottom,
    BehindText,
    InFrontOfText,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct OelDrawing {
    pub id: String,
    pub width_pt: f32,
    pub height_pt: f32,
    pub is_floating: bool,
    pub offset_x_pt: f32,
    pub offset_y_pt: f32,
    pub wrapping_mode: OelWrappingMode,
}

/// A text run: contiguous text with uniform formatting.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelRun {
    pub text: String,
    pub drawing: Option<OelDrawing>,
    pub props: OelRunProps,
}

impl OelRun {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            text: text.into(),
            drawing: None,
            props: OelRunProps::default(),
        }
    }

    pub fn with_props(text: impl Into<String>, props: OelRunProps) -> Self {
        Self {
            text: text.into(),
            drawing: None,
            props,
        }
    }

    pub fn with_drawing(drawing: OelDrawing, props: OelRunProps) -> Self {
        Self {
            text: String::new(),
            drawing: Some(drawing),
            props,
        }
    }

    /// Number of Unicode scalar values (chars) in this run.
    pub fn char_len(&self) -> usize {
        self.text.chars().count() + if self.drawing.is_some() { 1 } else { 0 }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelParagraph {
    /// Stable ID for React rendering keys.
    pub id: String,
    pub props: OelParaProps,
    pub runs: Vec<OelRun>,
}

impl OelParagraph {
    pub fn new(id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            props: OelParaProps::default(),
            runs: Vec::new(),
        }
    }

    pub fn with_text(id: impl Into<String>, text: impl Into<String>) -> Self {
        let mut p = Self::new(id);
        p.runs.push(OelRun::new(text));
        p
    }

    /// Total character count across all runs.
    pub fn char_len(&self) -> usize {
        self.runs.iter().map(|r| r.char_len()).sum()
    }

    /// Collect all text as a single String.
    pub fn plain_text(&self) -> String {
        self.runs.iter().map(|r| r.text.as_str()).collect()
    }

    /// Merge adjacent runs that share identical props.
    pub fn normalize_runs(&mut self) {
        if self.runs.len() <= 1 {
            return;
        }
        let mut merged: Vec<OelRun> = Vec::new();
        for run in self.runs.drain(..) {
            if let Some(last) = merged.last_mut() {
                if last.props == run.props && last.drawing.is_none() && run.drawing.is_none() {
                    last.text.push_str(&run.text);
                    continue;
                }
            }
            merged.push(run);
        }
        self.runs = merged;
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelTableCell {
    pub blocks: Vec<OelBlock>,
    pub props: OelTableCellProps,
}

impl OelTableCell {
    pub fn new() -> Self {
        let mut cell = Self {
            blocks: Vec::new(),
            props: OelTableCellProps::default(),
        };
        cell.blocks
            .push(OelBlock::Paragraph(OelParagraph::new(next_id())));
        cell
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelTableRow {
    pub cells: Vec<OelTableCell>,
}

impl OelTableRow {
    pub fn new(cols: usize) -> Self {
        Self {
            cells: (0..cols).map(|_| OelTableCell::new()).collect(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelTable {
    pub id: String,
    pub rows: Vec<OelTableRow>,
    pub props: OelTableProps,
}

impl OelTable {
    pub fn new(id: impl Into<String>, rows: usize, cols: usize) -> Self {
        Self {
            id: id.into(),
            rows: (0..rows).map(|_| OelTableRow::new(cols)).collect(),
            props: OelTableProps::default(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", content = "data")]
pub enum OelBlock {
    Paragraph(OelParagraph),
    Table(OelTable),
    PageBreak,
}

impl OelBlock {
    pub fn as_paragraph(&self) -> Option<&OelParagraph> {
        match self {
            OelBlock::Paragraph(p) => Some(p),
            _ => None,
        }
    }

    pub fn as_paragraph_mut(&mut self) -> Option<&mut OelParagraph> {
        match self {
            OelBlock::Paragraph(p) => Some(p),
            _ => None,
        }
    }
}

/// Simple monotonic counter for stable block IDs.
/// Works in both native and WASM (single-threaded).
pub fn next_id() -> String {
    use std::cell::Cell;
    thread_local! {
        static COUNTER: Cell<u64> = const { Cell::new(1) };
    }
    let id = COUNTER.with(|c| {
        let v = c.get();
        c.set(v + 1);
        v
    });
    format!("b{id}")
}
