# oeldocx-rs

`oeldocx-rs` is a WebAssembly-based DOCX document editing engine built for the OutilsEnLigne document editor. It provides a robust, memory-efficient, and type-safe core for parsing, editing, and serializing Word documents (`.docx`) natively in the browser.

## Features

- **WASM-Powered Performance:** Runs entirely in the browser using WebAssembly.
- **DOCX Parsing & Serialization:** Leverages the [`docx-rs`](https://github.com/bokuweb/docx-rs) crate for reading and writing standard `.docx` files.
- **Stateful Editing Engine:** Maintains a mutable in-memory document model (`OelDocument`) handling text, paragraphs, lists, and tables.
- **Cursor & Selection Management:** Tracks caret position and text selection across standard blocks and table cells.
- **Rich Text Formatting:** Supports bold, italic, underline, strikethrough, font family/size, colors, text alignment, and paragraph spacing.
- **Undo/Redo History:** Built-in history stack to effortlessly revert or reapply changes.
- **Type-Safe WASM Bridge:** Seamless serialization of the editor state to JavaScript via `serde-wasm-bindgen`.

## Architecture

The frontend (React + Vite) uses `oeldocx-rs` as a headless stateful controller. 

1. **Rust (`DocxController`)**: The single source of truth for the document's state. It receives commands (e.g., `insert_text`, `set_bold`, `insert_table`), mutates the internal `OelDocument`, and returns a serialized `EditorState` to the UI.
2. **WASM Boundary (`src/wasm/bindings.rs`)**: Exposes the `DocxController` to JavaScript. Methods are explicitly bound using `wasm-bindgen`.
3. **TypeScript Frontend**: Calls the WASM methods and uses the returned `EditorState` to render the document on the canvas (handled by the Vite frontend repository).

## Directory Structure

- `src/model/`: In-memory data types representing the document (paragraphs, runs, tables, properties).
- `src/controller/`: The core `DocxController` engine logic and state orchestration.
- `src/commands/`: Granular document mutation commands (text insertion, table manipulation, formatting).
- `src/convert/`: Conversion layers bridging external `docx-rs` structures and the internal `OelDocument`.
- `src/state/`: Structs representing the `EditorState` which gets exported to the frontend.
- `src/render/`: Layout/Rendering structures optimized for UI consumption.
- `src/wasm/`: WebAssembly binding definitions.
- `src/cursor/` & `src/history/`: Cursor selection logic and Undo/Redo stack implementations.

## Building

To build the WASM package for the frontend, you will need standard Rust tooling along with `wasm-pack`. (Typically, the root workspace or frontend repo will have scripts that orchestrate this).

```bash
# Install wasm-pack if you haven't already
cargo install wasm-pack

# Build the package for web consumption
wasm-pack build --target web --release
```

*Note: The generated WASM binary and JS glue code will be output to the `pkg/` directory.*

## Documentation & Planning

For a deep dive into the underlying data structures, design decisions, and architectural specifications, please refer to the detailed [`PLAN.md`](./PLAN.md) file included in this repository.