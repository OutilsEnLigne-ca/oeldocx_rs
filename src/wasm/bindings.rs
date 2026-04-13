use crate::controller::DocxController as InnerController;
use crate::state::EditorState;
use wasm_bindgen::prelude::*;

fn to_js(state: &EditorState) -> Result<JsValue, JsValue> {
    serde_wasm_bindgen::to_value(state).map_err(|e| JsValue::from_str(&e.to_string()))
}

fn err_js(msg: impl std::fmt::Display) -> JsValue {
    JsValue::from_str(&msg.to_string())
}

/// WASM-exposed DOCX editing controller.
///
/// Every mutating method returns the new `EditorState` as a JS object.
/// Call `get_state()` at any time to read the current state without mutating.
#[wasm_bindgen]
pub struct JsExtractedFont {
    name: String,
    style: String,
    data: Vec<u8>,
}

#[wasm_bindgen]
impl JsExtractedFont {
    #[wasm_bindgen(getter)]
    pub fn name(&self) -> String {
        self.name.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn style(&self) -> String {
        self.style.clone()
    }

    #[wasm_bindgen(getter)]
    pub fn data(&self) -> js_sys::Uint8Array {
        js_sys::Uint8Array::from(self.data.as_slice())
    }
}

/// Extract embedded fonts from DOCX bytes without loading the full document.
#[wasm_bindgen]
pub fn extract_embedded_fonts(bytes: &[u8]) -> Result<js_sys::Array, JsValue> {
    crate::fonts::extract_fonts(bytes)
        .map(|fonts| {
            let arr = js_sys::Array::new();
            for f in fonts {
                let js_font = JsExtractedFont {
                    name: f.name,
                    style: f.style,
                    data: f.data,
                };
                arr.push(&JsValue::from(js_font));
            }
            arr
        })
        .map_err(|e| err_js(e))
}

#[wasm_bindgen]
pub struct DocxController {
    inner: InnerController,
}

#[wasm_bindgen]
impl DocxController {
    #[wasm_bindgen(constructor)]
    pub fn new() -> DocxController {
        console_error_panic_hook::set_once();
        DocxController {
            inner: InnerController::new(),
        }
    }

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    /// Load a DOCX file from raw bytes. Returns the initial `EditorState`.
    pub fn load(&mut self, bytes: &[u8], filename: Option<String>) -> Result<JsValue, JsValue> {
        self.inner
            .load(bytes, filename)
            .map_err(|e| err_js(e))
            .and_then(|s| to_js(&s))
    }

    /// Start a fresh empty document.
    pub fn new_document(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.new_document())
    }

    /// Serialize the current document to DOCX bytes.
    pub fn serialize(&self) -> Result<Box<[u8]>, JsValue> {
        self.inner
            .serialize()
            .map(|v| v.into_boxed_slice())
            .map_err(|e| err_js(e))
    }

    /// Read the current `EditorState` without mutating.
    pub fn get_state(&self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.get_state())
    }

    /// Get all extracted images as a record of ID to Base64 data URL.
    pub fn get_images(&self) -> Result<JsValue, JsValue> {
        use serde::Serialize;
        let serializer = serde_wasm_bindgen::Serializer::new().serialize_maps_as_objects(true);
        self.inner.images.serialize(&serializer).map_err(|e| err_js(e))
    }

    // -------------------------------------------------------------------------
    // Cursor / Selection
    // -------------------------------------------------------------------------

    /// Sync the cursor/selection from the JS renderer.
    /// Pass `anchor_row = -1` (as i32) to indicate not in a table.
    #[wasm_bindgen]
    pub fn set_selection(
        &mut self,
        anchor_block: usize,
        anchor_row: i32,
        anchor_col: i32,
        anchor_inner: usize,
        anchor_offset: usize,
        focus_block: usize,
        focus_row: i32,
        focus_col: i32,
        focus_inner: usize,
        focus_offset: usize,
    ) -> Result<JsValue, JsValue> {
        let opt = |v: i32| if v < 0 { None } else { Some(v as usize) };
        let s = self.inner.set_selection(
            anchor_block,
            opt(anchor_row),
            opt(anchor_col),
            anchor_inner,
            anchor_offset,
            focus_block,
            opt(focus_row),
            opt(focus_col),
            focus_inner,
            focus_offset,
        );
        to_js(&s)
    }

    // -------------------------------------------------------------------------
    // Text input
    // -------------------------------------------------------------------------

    pub fn insert_text(&mut self, text: &str) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_text(text))
    }

    pub fn delete_backward(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.delete_backward())
    }

    pub fn delete_forward(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.delete_forward())
    }

    pub fn insert_newline(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_newline())
    }

    // -------------------------------------------------------------------------
    // Character formatting
    // -------------------------------------------------------------------------

    pub fn set_bold(&mut self, value: bool) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_bold(value))
    }

    pub fn set_italic(&mut self, value: bool) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_italic(value))
    }

    pub fn set_underline(&mut self, value: bool) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_underline(value))
    }

    pub fn set_strikethrough(&mut self, value: bool) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_strikethrough(value))
    }

    /// `half_points`: font size in OOXML half-points (e.g. 24 = 12pt).
    pub fn set_font_size(&mut self, half_points: u32) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_font_size(half_points))
    }

    pub fn set_font_family(&mut self, family: &str) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_font_family(family))
    }

    pub fn set_text_color(&mut self, hex: &str) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_text_color(hex))
    }

    pub fn set_highlight(&mut self, hex: Option<String>) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_highlight(hex.as_deref()))
    }

    // -------------------------------------------------------------------------
    // Paragraph formatting
    // -------------------------------------------------------------------------

    pub fn set_alignment_left(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_alignment(crate::model::Alignment::Left))
    }
    pub fn set_alignment_center(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_alignment(crate::model::Alignment::Center))
    }
    pub fn set_alignment_right(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_alignment(crate::model::Alignment::Right))
    }
    pub fn set_alignment_justify(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_alignment(crate::model::Alignment::Justify))
    }

    pub fn set_indent(&mut self, level: u32) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_indent(level))
    }

    pub fn increase_indent(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.increase_indent())
    }

    pub fn decrease_indent(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.decrease_indent())
    }

    pub fn toggle_bullet_list(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.toggle_bullet_list())
    }

    pub fn toggle_numbered_list(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.toggle_numbered_list())
    }

    pub fn set_line_spacing(&mut self, multiplier: f32) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_line_spacing(multiplier))
    }

    pub fn set_paragraph_spacing(&mut self, before: u32, after: u32) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_paragraph_spacing(before, after))
    }

    pub fn apply_style(&mut self, style_id: &str) -> Result<JsValue, JsValue> {
        to_js(&self.inner.apply_style(style_id))
    }

    // -------------------------------------------------------------------------
    // Table operations
    // -------------------------------------------------------------------------

    pub fn insert_table(&mut self, rows: u32, cols: u32) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_table(rows, cols))
    }

    pub fn insert_row_above(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_row_above())
    }

    pub fn insert_row_below(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_row_below())
    }

    pub fn insert_col_left(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_col_left())
    }

    pub fn insert_col_right(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_col_right())
    }

    pub fn delete_row(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.delete_row())
    }

    pub fn delete_col(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.delete_col())
    }

    pub fn delete_table(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.delete_table())
    }

    // -------------------------------------------------------------------------
    // Section / page
    // -------------------------------------------------------------------------

    pub fn insert_page_break(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_page_break())
    }

    /// Width and height in twips (1/1440 inch).
    pub fn set_page_size(
        &mut self,
        width_twips: u32,
        height_twips: u32,
    ) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_page_size(width_twips, height_twips))
    }

    pub fn set_margins(
        &mut self,
        top: u32,
        right: u32,
        bottom: u32,
        left: u32,
    ) -> Result<JsValue, JsValue> {
        to_js(&self.inner.set_margins(top, right, bottom, left))
    }

    // -------------------------------------------------------------------------
    // Images
    // -------------------------------------------------------------------------

    /// Insert a PNG image at the current cursor position as an inline drawing.
    ///
    /// After this call, invoke `get_images()` to retrieve the new image's data URL.
    pub fn insert_image(
        &mut self,
        data: &[u8],
        width_pt: f32,
        height_pt: f32,
    ) -> Result<JsValue, JsValue> {
        to_js(&self.inner.insert_image(data.to_vec(), width_pt, height_pt))
    }

    /// Change the wrapping mode of an image.
    ///
    /// `mode` is a camelCase string matching the `OelWrappingMode` enum:
    /// `"inline"`, `"square"`, `"tight"`, `"through"`, `"topAndBottom"`,
    /// `"behindText"`, `"inFrontOfText"`.
    pub fn update_image_wrap(
        &mut self,
        image_id: String,
        mode: String,
    ) -> Result<JsValue, JsValue> {
        let wrapping_mode: crate::model::block::OelWrappingMode =
            serde_json::from_str(&format!("\"{}\"", mode))
                .map_err(|e| err_js(e))?;
        to_js(&self.inner.update_image_wrap(&image_id, wrapping_mode))
    }

    /// Reposition a floating image. `x_pt` and `y_pt` are in points relative to the anchor.
    pub fn move_image(
        &mut self,
        image_id: String,
        x_pt: f32,
        y_pt: f32,
    ) -> Result<JsValue, JsValue> {
        to_js(&self.inner.move_image(&image_id, x_pt, y_pt))
    }

    /// Resize an image to the given width and height in points.
    pub fn resize_image(
        &mut self,
        image_id: String,
        width_pt: f32,
        height_pt: f32,
    ) -> Result<JsValue, JsValue> {
        to_js(&self.inner.resize_image(&image_id, width_pt, height_pt))
    }

    // -------------------------------------------------------------------------
    // History
    // -------------------------------------------------------------------------

    pub fn undo(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.undo())
    }

    pub fn redo(&mut self) -> Result<JsValue, JsValue> {
        to_js(&self.inner.redo())
    }

    // -------------------------------------------------------------------------
    // Clipboard
    // -------------------------------------------------------------------------

    pub fn copy(&self) -> String {
        self.inner.copy()
    }

    pub fn paste(&mut self, text: &str) -> Result<JsValue, JsValue> {
        to_js(&self.inner.paste(text))
    }
}
