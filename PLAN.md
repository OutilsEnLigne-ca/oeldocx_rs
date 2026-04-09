# oeldocx_rs — Architecture Plan

> Living document. Update as decisions are made. Written to hand to a future Claude Code instance.

---

## What this crate is

`oeldocx_rs` is a Rust/WASM document editing engine for DOCX files.
It sits on top of `docx-rs` (parse + serialize) and adds everything needed for interactive editing:
an in-memory document AST, cursor/selection model, command layer, render tree generation,
undo/redo history, and WASM bindings.

`docx-rs` knows the DOCX format.
`oeldocx_rs` knows how to edit a document.

---

## What `docx-rs` provides (the base interface)

`docx-rs` 0.4.x gives us:

- `read_docx(bytes: &[u8]) -> Result<DocxDocument, ReaderError>` — parses a `.docx` archive
- `Docx` builder for constructing documents (paragraphs, runs, tables, sections)
- Parsed types: `DocumentChild::Paragraph`, `DocumentChild::Table`, each with nested children
- `RunProperty` — bold, italic, underline, strike, font size, font name, color, highlight
- `ParagraphProperty` — alignment, indent, spacing, numbering, style
- `TableProperty`, `TableRowProperty`, `TableCellProperty`
- `SectionProperty` — page size, margins, headers, footers

What `docx-rs` does NOT provide:
- Cursor or selection tracking
- Mutation commands (it is a builder, not a live editor)
- Render output (it serializes to OOXML bytes, not a render tree)
- Undo/redo
- WASM-friendly API

---

## Extension strategy

We do NOT mutate `docx-rs` types directly.
Instead, we define our own **`OelDocument` AST** as the mutable in-memory model.

```
DOCX bytes  ──[docx-rs read_docx]──>  DocxDocument
                                           │
                                   [from_docx converter]
                                           │
                                           v
                                     OelDocument  ◄──── cursor, history
                                           │
                          ┌────────────────┴──────────────────┐
                          │                                   │
                  [to_render converter]             [to_docx converter]
                          │                                   │
                          v                                   v
                   RenderDocument                      docx-rs Docx
                 (sent to React via WASM)           ──[build()]──> bytes
```

`docx-rs` is used only at the two boundaries (parse in, serialize out).
All editing logic operates on `OelDocument`.

---

## Crate layout

```
oeldocx_rs/
  Cargo.toml
  src/
    lib.rs                  ← crate root; wasm-bindgen re-exports
    model/
      mod.rs
      document.rs           ← OelDocument: root AST node
      block.rs              ← OelBlock enum (paragraph, table, page-break)
      inline.rs             ← OelInline enum (run, image-inline, hyperlink)
      style.rs              ← OelRunProps, OelParaProps, formatting types
      section.rs            ← OelSectionProps (page size, margins, header/footer refs)
    cursor/
      mod.rs
      position.rs           ← DocPosition: address of a character in the document
      selection.rs          ← DocSelection: anchor + focus DocPosition
    history/
      mod.rs
      snapshot.rs           ← Snapshot = cloned OelDocument; UndoStack<Snapshot>
    convert/
      mod.rs
      from_docx.rs          ← DocxDocument → OelDocument
      to_docx.rs            ← OelDocument → docx-rs Docx
      to_render.rs          ← OelDocument + DocSelection → RenderDocument
    render/
      mod.rs
      types.rs              ← RenderDocument, RenderBlock, RenderSpan, RenderSectionProps, etc.
    commands/
      mod.rs
      text.rs               ← insert_text, delete_backward, delete_forward
      format.rs             ← set_bold, set_italic, set_font_size, set_color, etc.
      paragraph.rs          ← set_alignment, set_indent, toggle_list
      table.rs              ← insert_table, insert/delete row/col
      section.rs            ← set_page_size, set_margins, insert_page_break
    state/
      mod.rs
      types.rs              ← EditorState, FormatState, CursorState, SelectionState,
                               HistoryState, DocumentInfo
    controller/
      mod.rs
      docx_controller.rs    ← DocxController: the main stateful object
    wasm/
      mod.rs
      bindings.rs           ← #[wasm_bindgen] wrapper around DocxController
```

---

## Data types

### `OelDocument` — the mutable in-memory model

```rust
pub struct OelDocument {
    pub blocks: Vec<OelBlock>,
    pub section: OelSectionProps,
    pub styles: HashMap<String, OelStyle>,  // named paragraph styles
}

pub enum OelBlock {
    Paragraph(OelParagraph),
    Table(OelTable),
    PageBreak,
}

pub struct OelParagraph {
    pub id: String,               // stable UUID for React keys
    pub props: OelParaProps,
    pub runs: Vec<OelRun>,
}

pub struct OelRun {
    pub text: String,             // raw text content
    pub props: OelRunProps,
}

pub struct OelRunProps {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size: Option<u32>,   // half-points (OOXML convention), None = inherit
    pub font_family: Option<String>,
    pub color: Option<String>,    // hex RGB, None = inherit
    pub highlight: Option<String>,
}

pub struct OelParaProps {
    pub alignment: Alignment,     // Left | Center | Right | Justify
    pub indent_level: u32,
    pub list_type: Option<ListType>,  // Bullet | Numbered(style)
    pub spacing_before: Option<u32>,  // pt
    pub spacing_after: Option<u32>,
    pub line_spacing: Option<f32>,
    pub style_id: Option<String>,
}

pub struct OelTable {
    pub id: String,
    pub rows: Vec<OelTableRow>,
    pub props: OelTableProps,
}

pub struct OelTableRow {
    pub cells: Vec<OelTableCell>,
}

pub struct OelTableCell {
    pub blocks: Vec<OelBlock>,   // cell contains its own block list
    pub props: OelTableCellProps,
}

pub struct OelSectionProps {
    pub page_width: u32,   // twips (1/1440 inch)
    pub page_height: u32,
    pub margin_top: u32,
    pub margin_right: u32,
    pub margin_bottom: u32,
    pub margin_left: u32,
}
```

### `DocPosition` — cursor addressing

Cursors address characters using a two-level path.
Phase 1 supports top-level paragraphs and table cells.

```rust
pub struct DocPosition {
    pub block_idx: usize,
    pub cell_path: Option<(usize, usize)>,  // (row, col) if inside a table cell
    pub inner_block_idx: usize,             // block within cell (0 for top-level paras)
    pub run_idx: usize,
    pub char_offset: usize,                 // char index (not byte), within run.text
}

pub struct DocSelection {
    pub anchor: DocPosition,  // where selection started
    pub focus: DocPosition,   // where cursor currently is (may be before anchor)
}

impl DocSelection {
    pub fn is_collapsed(&self) -> bool;       // anchor == focus (no range selected)
    pub fn ordered(&self) -> (&DocPosition, &DocPosition);  // (start, end) in doc order
}
```

### `EditorState` — the WASM boundary type

Everything the frontend needs after each command. Serialized as JSON via `serde`.

```rust
#[derive(Serialize, Deserialize)]
pub struct EditorState {
    pub document: RenderDocument,
    pub cursor: CursorState,
    pub selection: SelectionState,
    pub format: FormatState,     // formatting at cursor (read from run under cursor)
    pub history: HistoryState,
    pub document_info: DocumentInfo,
}

#[derive(Serialize, Deserialize)]
pub struct FormatState {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size: u32,           // resolved (inheriting from default style if None)
    pub font_family: String,
    pub color: String,            // hex
    pub highlight: Option<String>,
    pub alignment: Alignment,
    pub list_type: Option<ListType>,
    pub is_in_table: bool,
}

#[derive(Serialize, Deserialize)]
pub struct HistoryState {
    pub has_undo: bool,
    pub has_redo: bool,
}

#[derive(Serialize, Deserialize)]
pub struct DocumentInfo {
    pub word_count: usize,
    pub char_count: usize,
    pub filename: Option<String>,
}
```

---

## `RenderDocument` — the render tree

The WASM layer does not paginate. The render tree is a flat block list.
The React renderer simulates page boundaries with CSS (page-break rules, fixed-height containers).
Real pagination (line-wrapping, exact page breaks) is deferred to a later phase.

```rust
#[derive(Serialize, Deserialize)]
pub struct RenderDocument {
    pub blocks: Vec<RenderBlock>,
    pub section: RenderSectionProps,
}

#[derive(Serialize, Deserialize)]
pub struct RenderSectionProps {
    pub page_width_pt: f32,
    pub page_height_pt: f32,
    pub margin_top_pt: f32,
    pub margin_right_pt: f32,
    pub margin_bottom_pt: f32,
    pub margin_left_pt: f32,
}

#[derive(Serialize, Deserialize)]
pub enum RenderBlock {
    Paragraph(RenderParagraph),
    Table(RenderTable),
    PageBreak,
}

#[derive(Serialize, Deserialize)]
pub struct RenderParagraph {
    pub id: String,
    pub alignment: Alignment,
    pub indent_level: u32,
    pub list_type: Option<ListType>,
    pub list_index: Option<u32>,        // resolved counter for numbered lists
    pub spacing_before_pt: f32,
    pub spacing_after_pt: f32,
    pub spans: Vec<RenderSpan>,
}

#[derive(Serialize, Deserialize)]
pub struct RenderSpan {
    pub text: String,
    pub format: RenderFormat,
    // Char positions relative to start of this paragraph (for cursor overlay)
    pub char_start: usize,
    pub char_end: usize,
}

#[derive(Serialize, Deserialize)]
pub struct RenderFormat {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size_pt: f32,
    pub font_family: String,
    pub color: String,
    pub highlight: Option<String>,
}

#[derive(Serialize, Deserialize)]
pub struct RenderTable {
    pub id: String,
    pub rows: Vec<RenderTableRow>,
}

#[derive(Serialize, Deserialize)]
pub struct RenderTableRow {
    pub cells: Vec<RenderTableCell>,
}

#[derive(Serialize, Deserialize)]
pub struct RenderTableCell {
    pub blocks: Vec<RenderBlock>,
    pub col_span: u32,
    pub row_span: u32,
}
```

---

## `DocxController` — the stateful editing engine

```rust
pub struct DocxController {
    document: OelDocument,
    selection: DocSelection,
    undo_stack: UndoStack,
    filename: Option<String>,
}
```

### Command execution pattern

Every command follows this pattern:

```rust
fn exec_command(&mut self, apply: impl FnOnce(&mut OelDocument, &mut DocSelection)) -> EditorState {
    self.undo_stack.push_snapshot(&self.document);
    apply(&mut self.document, &mut self.selection);
    self.build_state()
}

fn build_state(&self) -> EditorState {
    EditorState {
        document: to_render::convert(&self.document, &self.selection),
        cursor: self.cursor_state(),
        selection: self.selection_state(),
        format: self.format_at_cursor(),
        history: self.history_state(),
        document_info: self.document_info(),
    }
}
```

High-frequency commands (single character insert/delete) skip snapshotting to avoid
O(n) clones per keystroke. Instead they accumulate into a single undo batch:

```rust
fn exec_text_input(&mut self, apply: impl FnOnce(&mut OelDocument, &mut DocSelection)) -> EditorState {
    // Push snapshot only if last command was NOT a text input
    if !self.last_was_text_input {
        self.undo_stack.push_snapshot(&self.document);
    }
    self.last_was_text_input = true;
    apply(&mut self.document, &mut self.selection);
    self.build_state()
}
```

### Full method surface

```rust
impl DocxController {
    pub fn new() -> Self;

    // Lifecycle
    pub fn load(&mut self, bytes: &[u8], filename: Option<String>) -> Result<EditorState, EditorError>;
    pub fn new_document(&mut self) -> EditorState;
    pub fn serialize(&self) -> Result<Vec<u8>, EditorError>;

    // Cursor / selection  (JS sends these BEFORE formatting commands)
    pub fn set_selection(&mut self, anchor: DocPosition, focus: DocPosition) -> EditorState;
    pub fn move_cursor(&mut self, direction: Direction, extend_selection: bool) -> EditorState;

    // Text input
    pub fn insert_text(&mut self, text: &str) -> EditorState;
    pub fn delete_backward(&mut self) -> EditorState;
    pub fn delete_forward(&mut self) -> EditorState;
    pub fn insert_newline(&mut self) -> EditorState;

    // Character formatting (apply to selection, or toggle at cursor)
    pub fn set_bold(&mut self, value: bool) -> EditorState;
    pub fn set_italic(&mut self, value: bool) -> EditorState;
    pub fn set_underline(&mut self, value: bool) -> EditorState;
    pub fn set_strikethrough(&mut self, value: bool) -> EditorState;
    pub fn set_font_size(&mut self, size: u32) -> EditorState;
    pub fn set_font_family(&mut self, name: &str) -> EditorState;
    pub fn set_text_color(&mut self, hex: &str) -> EditorState;
    pub fn set_highlight(&mut self, hex: Option<&str>) -> EditorState;

    // Paragraph formatting
    pub fn set_alignment(&mut self, alignment: Alignment) -> EditorState;
    pub fn set_indent(&mut self, level: u32) -> EditorState;
    pub fn toggle_bullet_list(&mut self) -> EditorState;
    pub fn toggle_numbered_list(&mut self) -> EditorState;
    pub fn set_line_spacing(&mut self, value: f32) -> EditorState;
    pub fn set_paragraph_spacing(&mut self, before: u32, after: u32) -> EditorState;
    pub fn apply_style(&mut self, style_id: &str) -> EditorState;

    // Table operations
    pub fn insert_table(&mut self, rows: u32, cols: u32) -> EditorState;
    pub fn insert_row_above(&mut self) -> EditorState;
    pub fn insert_row_below(&mut self) -> EditorState;
    pub fn insert_col_left(&mut self) -> EditorState;
    pub fn insert_col_right(&mut self) -> EditorState;
    pub fn delete_row(&mut self) -> EditorState;
    pub fn delete_col(&mut self) -> EditorState;
    pub fn delete_table(&mut self) -> EditorState;

    // Section / page
    pub fn insert_page_break(&mut self) -> EditorState;
    pub fn set_page_size(&mut self, width_twips: u32, height_twips: u32) -> EditorState;
    pub fn set_margins(&mut self, top: u32, right: u32, bottom: u32, left: u32) -> EditorState;

    // History
    pub fn undo(&mut self) -> EditorState;
    pub fn redo(&mut self) -> EditorState;

    // Clipboard (JS manages the system clipboard; Rust manages the internal payload)
    pub fn copy(&self) -> ClipboardPayload;
    pub fn cut(&mut self) -> (EditorState, ClipboardPayload);
    pub fn paste(&mut self, payload: ClipboardPayload) -> EditorState;

    // State query
    pub fn get_state(&self) -> EditorState;
}
```

---

## `UndoStack`

```rust
pub struct UndoStack {
    undo: Vec<OelDocument>,   // stack of past states
    redo: Vec<OelDocument>,   // stack of future states (cleared on new command)
    max_depth: usize,         // default: 100
}

impl UndoStack {
    pub fn push_snapshot(&mut self, doc: &OelDocument);  // clone doc, clear redo
    pub fn undo(&mut self, current: OelDocument) -> Option<OelDocument>;
    pub fn redo(&mut self, current: OelDocument) -> Option<OelDocument>;
    pub fn has_undo(&self) -> bool;
    pub fn has_redo(&self) -> bool;
}
```

`OelDocument` must derive `Clone`. The clone cost is proportional to document size.
For Phase 1 this is acceptable. For large documents, consider structural sharing (Rc/Arc on runs).

---

## `convert/from_docx.rs` — mapping from `docx-rs`

`docx-rs` parsed types → `OelDocument`. Key mappings:

| docx-rs type | Our type |
|---|---|
| `DocumentChild::Paragraph` | `OelBlock::Paragraph` |
| `DocumentChild::Table` | `OelBlock::Table` |
| `Run` with `RunChild::Text` | `OelRun { text, props }` |
| `RunProperty.bold` | `OelRunProps.bold` |
| `RunProperty.color` | `OelRunProps.color` (extract hex from `Color::Hex`) |
| `RunProperty.sz` | `OelRunProps.font_size` (half-points) |
| `ParagraphProperty.alignment` | `OelParaProps.alignment` |
| `NumberingProperty` | `OelParaProps.list_type + indent_level` |
| `SectionProperty` (last paragraph) | `OelSectionProps` |

Runs that are adjacent and share identical `OelRunProps` are merged during import
to keep the run list compact and avoid excessive spans in the render tree.

---

## `convert/to_docx.rs` — mapping to `docx-rs`

`OelDocument` → rebuild a `docx-rs` `Docx` using its builder API, then `.pack(cursor)` to bytes.

Key challenges:
- Numbered list numbering must be expressed via `docx-rs` `AbstractNumbering` and `Numbering` (OOXML abstract num IDs). We generate these fresh on each serialize.
- Paragraph styles must be re-emitted in the `Styles` part.
- Section properties must be attached to the last paragraph of the section.

---

## `convert/to_render.rs`

Walk `OelDocument`, produce `RenderDocument`. Resolves:
- Inherited formatting (font size, color, font family fall back through style chain → document defaults)
- List counters (number each `OelBlock::Paragraph` with `list_type: Numbered` sequentially)
- `char_start` / `char_end` per span (accumulate per paragraph for cursor overlay)

Selection is NOT embedded in the render tree. The React renderer positions the cursor overlay
by block ID and char offset, which it gets from `EditorState.cursor` / `.selection`.

---

## WASM bindings (`wasm/bindings.rs`)

Follows the `oelimg-rs` pattern closely.

```rust
#[wasm_bindgen]
pub struct DocxController {
    inner: controller::DocxController,
}

#[wasm_bindgen]
impl DocxController {
    #[wasm_bindgen(constructor)]
    pub fn new() -> DocxController { ... }

    pub fn load(&mut self, bytes: &[u8], filename: Option<String>) -> Result<JsValue, JsValue> {
        self.inner.load(bytes, filename)
            .map(|s| serde_wasm_bindgen::to_value(&s).unwrap())
            .map_err(|e| JsValue::from_str(&e.to_string()))
    }

    pub fn get_state(&self) -> JsValue {
        serde_wasm_bindgen::to_value(&self.inner.get_state()).unwrap()
    }

    // Every method: same Result<JsValue, JsValue> pattern
    pub fn set_bold(&mut self, value: bool) -> Result<JsValue, JsValue> { ... }
    pub fn insert_text(&mut self, text: &str) -> Result<JsValue, JsValue> { ... }
    pub fn set_selection(
        &mut self,
        anchor_block: usize, anchor_run: usize, anchor_offset: usize,
        focus_block: usize,  focus_run: usize,  focus_offset: usize,
    ) -> Result<JsValue, JsValue> { ... }
    // ... all commands forwarded
}
```

`set_selection` takes flat args (not a struct) to avoid needing a `JsValue` deserialize for a simple call.

---

## `Cargo.toml`

```toml
[package]
name = "oeldocx_rs"
version = "0.1.0"
edition = "2024"

[lib]
crate-type = ["cdylib", "rlib"]

[dependencies]
wasm-bindgen = { version = "0.2", features = ["serde-serialize"] }
wasm-bindgen-futures = "0.4"
js-sys = "0.3"
web-sys = { version = "0.3", features = ["console"] }
serde = { version = "1", features = ["derive"] }
serde-wasm-bindgen = "0.6"
docx-rs = "0.4"
console_error_panic_hook = "0.1"
uuid = { version = "1", features = ["v4", "js"] }

[dev-dependencies]
# plain Rust tests (not WASM) for round-trip verification
wasm-bindgen-test = "0.3"

[profile.release]
opt-level = "s"
lto = true
```

---

## Build & link to oel-vite

```bash
# In oeldocx_rs/
wasm-pack build --target bundler --out-dir pkg

# In oel-vite/ package.json add:
"oeldocx_rs": "file:../oeldocx_rs/pkg"

npm install
```

Vite >= 5 handles WASM automatically. If needed, add `vite-plugin-wasm` and
`vite-plugin-top-level-await` to `vite.config.ts`.

---

## Implementation order

Follow this sequence strictly. Do not skip ahead.

1. **Read this file in full** before touching any code.
2. Set up `Cargo.toml` with all dependencies.
3. `model/` — define all structs (derive `Clone`, `Serialize`, `Deserialize` everywhere).
   Start with `OelRunProps`, `OelParaProps`, `OelRun`, `OelParagraph`, `OelBlock`, `OelDocument`.
4. `render/types.rs` — define `RenderDocument` and all render types (serde only, no logic yet).
5. `state/types.rs` — define `EditorState`, `FormatState`, etc.
6. `cursor/` — `DocPosition`, `DocSelection`.
7. `history/snapshot.rs` — `UndoStack`.
8. `convert/from_docx.rs` — parse a real DOCX file into `OelDocument`.
   Write a plain `#[test]` that loads a sample DOCX bytes and asserts paragraph count.
9. `convert/to_docx.rs` — `OelDocument` → bytes. Test round-trip: parse → serialize → parse again.
10. `convert/to_render.rs` — `OelDocument` → `RenderDocument`.
11. `commands/` — implement each command module.
12. `controller/docx_controller.rs` — wire everything together.
13. `wasm/bindings.rs` — expose to JS.
14. `wasm-pack build` and fix compile errors.
15. Link to `oel-vite` and write a minimal React test harness (load a DOCX, log the state).

---

## Open questions / decisions pending

- [ ] **Text input batching**: How long is a "text input batch" for undo? Options: (a) batch until non-text command, (b) batch with a time window. Recommendation: option (a) for simplicity.
- [ ] **docx-rs numbering API**: Abstract numbering in `docx-rs` is complex. For Phase 1, bullet lists may be emitted as simple `<w:numId>` references. Verify the exact API before implementing `to_docx.rs`.
- [ ] **Font size inheritance**: If `OelRunProps.font_size` is `None`, what is the default? Recommendation: resolve to 12pt (24 half-points) as a hard default, overridable by `OelStyle`.
- [ ] **Selection crossing table cells**: For Phase 1, forbid selections that cross cell boundaries. A selection is always within one cell or entirely outside all tables.
- [ ] **Image handling**: `docx-rs` embeds images as relationship parts. For Phase 1, images in a parsed DOCX are preserved as opaque blobs during round-trip but not rendered in the React renderer. Image insertion is a Phase 2 feature.
- [ ] **Header/footer editing**: Section properties include header/footer references. For Phase 1, preserve them during round-trip but do not expose editing commands.
