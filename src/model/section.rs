use serde::{Deserialize, Serialize};

/// Page and section layout properties. All measurements in twips (1/1440 inch).
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct OelSectionProps {
    pub page_width: u32,
    pub page_height: u32,
    pub margin_top: u32,
    pub margin_right: u32,
    pub margin_bottom: u32,
    pub margin_left: u32,
}

impl Default for OelSectionProps {
    fn default() -> Self {
        Self {
            // A4: 210mm × 297mm  →  11906 × 16838 twips
            page_width: 11906,
            page_height: 16838,
            // 1 inch margins all around
            margin_top: 1440,
            margin_right: 1440,
            margin_bottom: 1440,
            margin_left: 1440,
        }
    }
}

impl OelSectionProps {
    /// Content width in twips (page width minus left+right margins).
    pub fn content_width(&self) -> u32 {
        self.page_width
            .saturating_sub(self.margin_left)
            .saturating_sub(self.margin_right)
    }
}
