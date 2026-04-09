//! # oeldocx-rs
//!
//! A WebAssembly-based DOCX document editing engine for OutilsEnLigne.
//!
//! This crate provides an interactive editing controller (`DocxController`)
//! that can be compiled to WebAssembly. It handles document state, history
//! (undo/redo), cursor management, and formatting, while integrating with
//! `docx-rs` for serialization and deserialization of DOCX files.

pub mod commands;
pub mod controller;
pub mod convert;
pub mod cursor;
pub mod history;
pub mod model;
pub mod render;
pub mod state;
pub mod wasm;

pub use wasm::bindings::DocxController;
