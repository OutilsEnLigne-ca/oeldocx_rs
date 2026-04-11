use crate::model::{Alignment, ListType};
use serde::{Deserialize, Serialize};

/// Twips-to-points factor (1 pt = 20 twips).
pub const TWIPS_TO_PT: f32 = 0.05;
/// Half-points-to-points factor.
pub const HALFPT_TO_PT: f32 = 0.5;
/// Default resolved font size in points.
pub const DEFAULT_FONT_SIZE_PT: f32 = 11.0;
pub const DEFAULT_FONT_FAMILY: &str = "Roboto";
pub const DEFAULT_COLOR: &str = "000000";

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderSectionProps {
    pub page_width_pt: f32,
    pub page_height_pt: f32,
    pub margin_top_pt: f32,
    pub margin_right_pt: f32,
    pub margin_bottom_pt: f32,
    pub margin_left_pt: f32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderFormat {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size_pt: f32,
    pub font_family: String,
    /// Hex RGB without '#'.
    pub color: String,
    pub highlight: Option<String>,
}

impl Default for RenderFormat {
    fn default() -> Self {
        Self {
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_size_pt: DEFAULT_FONT_SIZE_PT,
            font_family: DEFAULT_FONT_FAMILY.to_string(),
            color: DEFAULT_COLOR.to_string(),
            highlight: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub enum RenderWrappingMode {
    Inline,
    Square,
    Tight,
    Through,
    TopAndBottom,
    BehindText,
    InFrontOfText,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderDrawing {
    pub id: String,
    pub width_pt: f32,
    pub height_pt: f32,
    pub is_floating: bool,
    pub offset_x_pt: f32,
    pub offset_y_pt: f32,
    pub wrapping_mode: RenderWrappingMode,
}

/// A contiguous span of text with uniform formatting.
/// `char_start` and `char_end` are relative to the start of the parent paragraph
/// and are used by the React renderer to position the cursor overlay.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderSpan {
    pub text: String,
    pub drawing: Option<RenderDrawing>,
    pub format: RenderFormat,
    pub char_start: usize,
    pub char_end: usize,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderParagraph {
    pub id: String,
    pub style_id: Option<String>,
    pub alignment: Alignment,
    pub indent_level: u32,
    pub list_type: Option<ListType>,
    /// The originating `w:numId`. Used by the frontend to group consecutive paragraphs
    /// into the same logical list (for counter restart detection).
    pub num_id: Option<u32>,
    /// Resolved counter for numbered lists (1-based), scoped per (num_id, level).
    /// None for bullets or non-list paragraphs.
    pub list_index: Option<u32>,
    pub spacing_before_pt: f32,
    pub spacing_after_pt: f32,
    /// Unitless CSS line-height multiplier (1.0 = single, 1.5 = one-and-a-half, 2.0 = double).
    /// None means use the renderer default.
    pub line_spacing: Option<f32>,
    pub spans: Vec<RenderSpan>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderTableCell {
    pub blocks: Vec<RenderBlock>,
    pub col_span: u32,
    pub row_span: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderTableRow {
    pub cells: Vec<RenderTableCell>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderTable {
    pub id: String,
    pub rows: Vec<RenderTableRow>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "type", content = "data")]
pub enum RenderBlock {
    Paragraph(RenderParagraph),
    Table(RenderTable),
    PageBreak,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RenderDocument {
    pub blocks: Vec<RenderBlock>,
    pub section: RenderSectionProps,
}
