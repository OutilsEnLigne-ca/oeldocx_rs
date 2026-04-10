# OutilsEnLigne DOCX Roadmap

This roadmap tracks the development of features required to bring OutilsEnLigne's DOCX capabilities on par with Microsoft Word, adhering to our Rust (WASM) / React (Vite) architecture.

## 🟢 Phase 1: Core Typography & The Color Picker

### 1.1 Text Alignment & Spacing
- [ ] **[Rust Parser]** Extract `<w:jc>` (Justification) and `<w:spacing>` from `w:pPr`.
- [ ] **[Rust Model]** Add `alignment` and `spacing` fields to `ParagraphStyle` in `OelDocument` and `RenderDocument`.
- [ ] **[React UI]** Map Rust `alignment` to CSS `text-align` and `spacing` to `margin`/`line-height`.
- [ ] **[Rust & React]** Implement `ApplyParagraphAlignmentCommand` and wire to UI buttons.

### 1.2 Custom Color Picker Component
- [ ] **[React UI]** Build `ColorPicker` component (Hex, RGB, Word's default palette).
- [ ] **[Rust Parser]** Parse `<w:color>` and `<w:highlight>` in `w:rPr`, normalizing Hex codes.
- [ ] **[Rust & React]** Implement `ApplyTextColorCommand` and `ApplyTextHighlightCommand`. Hook up to `ColorPicker`.

---

## 🟡 Phase 2: Lists, Indentation & Outlining

### 2.1 Numbering Parser Foundation
- [ ] **[Rust Parser]** Parse `word/numbering.xml`. Create `AbstractNum` and `NumId` lookup table.
- [ ] **[Rust Model]** Resolve prefixes (`1.`, `a.`, `•`) for paragraphs with `<w:numPr>`.
- [ ] **[Rust Engine]** Export resolved list types and levels via `RenderDocument`.

### 2.2 Frontend List Rendering
- [ ] **[React UI]** Map numbering data into semantic `<ul>`, `<ol>`, `<li>` or nested block components.
- [ ] **[Rust Commands]** Implement `ToggleListCommand` (Bullet/Numbered).
- [ ] **[Rust Commands]** Implement `ChangeListLevelCommand` (Tab/Shift+Tab) and handle new `NumId` generation.

---

## 🟠 Phase 3: Images & Object Layouts

### 3.1 Inline Images (Standard)
- [ ] **[Rust Parser]** Parse `DrawingML` (`<wp:inline>`). Base64 encode or extract image blobs.
- [ ] **[Rust Model]** Add `Image` node type to AST.
- [ ] **[React UI]** Render inline `<img src={...} />` inside paragraphs.

### 3.2 Floating Images & Text Wrapping
- [ ] **[Rust Parser]** Parse `<wp:anchor>` (wrapping styles, absolute coordinates).
- [ ] **[React UI]** Implement advanced CSS layout (absolute positioning, `float`, `clear`, `clip-path`).
- [ ] **[Rust & React]** Implement `InsertImageCommand`, `UpdateImageWrapCommand`, and `MoveImageCommand`.

---

## 🔵 Phase 4: Advanced Document Structures

### 4.1 Table of Contents (TOC)
- [ ] **[Rust Parser]** Detect TOC `w:sdt` field codes and read cached `<w:r>` runs.
- [ ] **[Rust Engine]** Dynamic TOC generation (AST traversal for Heading 1-9).
- [ ] **[React UI]** Render Document Outline sidebar using dynamic tree.
- [ ] **[Rust Commands]** Implement `UpdateTOCCommand`.

### 4.2 Complex Tables
- [ ] **[Rust Parser]** Parse `gridSpan` (colSpan), `vMerge` (rowSpan), borders, and shading in `w:tcPr`.
- [ ] **[React UI]** Render semantic HTML `<table>` with proper cell spans and borders.
- [ ] **[Rust Commands]** Implement `InsertTableRow`, `InsertTableCol`, `MergeCells`, `SplitCells`.

### 4.3 Headers, Footers & Sections
- [ ] **[Rust Parser]** Parse `word/header*.xml` and `word/footer*.xml`, mapping to `w:sectPr`.
- [ ] **[React UI]** Render visually separated headers and footers per section/page.

---

## 🟣 Phase 5: Fields, Interactivity & Optimization

### 5.1 Hyperlinks & Bookmarks
- [ ] **[Rust Parser]** Resolve `w:hyperlink` relationships and `w:bookmarkStart`/`End`.
- [ ] **[React UI]** Render clickable `<a>` tags and build link editing UI.

### 5.2 Performance & Web Worker Tuning
- [ ] **[Rust Engine]** Implement patch/diff-based updates for `RenderDocument` to avoid full serialization.
- [ ] **[TS/Web Worker]** Optimize state synchronization to maintain 60fps on large documents.
