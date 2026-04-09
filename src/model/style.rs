use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment::Left
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum ListType {
    Bullet,
    Numbered,
}

/// Formatting properties for a text run.
/// `None` on font_size / font_family / color means "inherit from paragraph style or defaults".
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OelRunProps {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    /// Half-points (OOXML convention). None = inherit.
    pub font_size: Option<u32>,
    /// None = inherit.
    pub font_family: Option<String>,
    /// Hex RGB without '#', e.g. "FF0000". None = inherit (black).
    pub color: Option<String>,
    /// Named highlight color or hex. None = no highlight.
    pub highlight: Option<String>,
}

impl Default for OelRunProps {
    fn default() -> Self {
        Self {
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_size: None,
            font_family: None,
            color: None,
            highlight: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OelParaProps {
    pub alignment: Alignment,
    pub indent_level: u32,
    pub list_type: Option<ListType>,
    /// Spacing before paragraph in twips.
    pub spacing_before: Option<u32>,
    /// Spacing after paragraph in twips.
    pub spacing_after: Option<u32>,
    /// Line spacing multiplier (1.0 = single, 2.0 = double).
    pub line_spacing: Option<f32>,
    /// Named style ID from the document's styles part.
    pub style_id: Option<String>,
}

impl Default for OelParaProps {
    fn default() -> Self {
        Self {
            alignment: Alignment::default(),
            indent_level: 0,
            list_type: None,
            spacing_before: None,
            spacing_after: None,
            line_spacing: None,
            style_id: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelTableProps {
    /// Width in twips. None = auto.
    pub width: Option<u32>,
}

impl Default for OelTableProps {
    fn default() -> Self {
        Self { width: None }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelTableCellProps {
    pub col_span: u32,
    pub row_span: u32,
    /// Background color hex. None = no fill.
    pub background: Option<String>,
}

impl Default for OelTableCellProps {
    fn default() -> Self {
        Self {
            col_span: 1,
            row_span: 1,
            background: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OelStyle {
    pub id: String,
    pub name: String,
    pub run_props: OelRunProps,
    pub para_props: OelParaProps,
}
