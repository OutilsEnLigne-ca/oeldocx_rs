use crate::model::{
    Alignment, ListType, OelBlock, OelDocument, OelParaProps, OelParagraph, OelRun, OelRunProps,
    OelSectionProps, OelStyle, OelTable, OelTableCell, OelTableCellProps, OelTableRow, next_id,
};
use docx_rs::{
    DocumentChild, Docx, ParagraphChild, RunChild, TableCellContent, TableChild, TableRowChild,
};
use std::collections::HashMap;

pub fn docx_to_oel(docx: &Docx) -> OelDocument {
    let blocks = docx
        .document
        .children
        .iter()
        .filter_map(convert_document_child)
        .collect();

    let section = convert_section(docx);

    let mut styles = OelDocument::empty().styles;

    if let Some(normal) = styles.get_mut("Normal") {
        if let Ok(json_val) = serde_json::to_value(&docx.styles.doc_defaults) {
            if let Some(rp) = json_val
                .get("runPropertyDefault")
                .and_then(|r| r.get("runProperty"))
            {
                if let Some(ff) = rp
                    .get("fonts")
                    .and_then(|f| f.get("ascii"))
                    .and_then(|a| a.as_str())
                {
                    normal.run_props.font_family = Some(ff.to_string());
                }
                if let Some(sz) = rp
                    .get("sz")
                    .and_then(|s| s.get("val"))
                    .and_then(|v| v.as_u64())
                {
                    normal.run_props.font_size = Some(sz as u32);
                }
                if let Some(color) = rp
                    .get("color")
                    .and_then(|c| c.get("val"))
                    .and_then(|v| v.as_str())
                {
                    if color != "auto" && !color.is_empty() {
                        normal.run_props.color = Some(color.to_string());
                    }
                }
            }
        }
    }

    for style in &docx.styles.styles {
        let run_props = convert_run_props(&style.run_property);
        let para_props = convert_para_props(&style.paragraph_property);

        let name = serde_str_field(&style.name, "val").unwrap_or_else(|| style.style_id.clone());

        styles.insert(
            style.style_id.clone(),
            OelStyle {
                id: style.style_id.clone(),
                name,
                run_props,
                para_props,
            },
        );
    }

    OelDocument {
        blocks,
        section,
        styles,
    }
}

fn convert_document_child(child: &DocumentChild) -> Option<OelBlock> {
    match child {
        DocumentChild::Paragraph(p) => {
            let mut para = OelParagraph::new(next_id());
            para.props = convert_para_props(&p.property);

            for pc in &p.children {
                if let ParagraphChild::Run(run) = pc {
                    let props = convert_run_props(&run.run_property);
                    let mut text = String::new();
                    for rc in &run.children {
                        match rc {
                            RunChild::Text(t) => text.push_str(&t.text),
                            RunChild::Tab(_) => text.push('\t'),
                            RunChild::Break(_) => text.push('\n'),
                            _ => {}
                        }
                    }
                    if !text.is_empty() {
                        para.runs.push(OelRun::with_props(text, props));
                    }
                }
            }
            para.normalize_runs();
            Some(OelBlock::Paragraph(para))
        }
        DocumentChild::Table(t) => {
            let id = next_id();
            let rows = t
                .rows
                .iter()
                .filter_map(|tc| {
                    let TableChild::TableRow(row) = tc;
                    let cells = row
                        .cells
                        .iter()
                        .filter_map(|rc| {
                            let TableRowChild::TableCell(cell) = rc;
                            let blocks = cell
                                .children
                                .iter()
                                .filter_map(|content| match content {
                                    TableCellContent::Paragraph(p) => {
                                        let mut para = OelParagraph::new(next_id());
                                        para.props = convert_para_props(&p.property);
                                        for pc in &p.children {
                                            if let ParagraphChild::Run(run) = pc {
                                                let props = convert_run_props(&run.run_property);
                                                let mut text = String::new();
                                                for rc in &run.children {
                                                    match rc {
                                                        RunChild::Text(t) => text.push_str(&t.text),
                                                        RunChild::Tab(_) => text.push('\t'),
                                                        _ => {}
                                                    }
                                                }
                                                if !text.is_empty() {
                                                    para.runs.push(OelRun::with_props(text, props));
                                                }
                                            }
                                        }
                                        para.normalize_runs();
                                        Some(OelBlock::Paragraph(para))
                                    }
                                    _ => None,
                                })
                                .collect::<Vec<_>>();

                            let blocks = if blocks.is_empty() {
                                vec![OelBlock::Paragraph(OelParagraph::new(next_id()))]
                            } else {
                                blocks
                            };

                            Some(OelTableCell {
                                blocks,
                                props: OelTableCellProps::default(),
                            })
                        })
                        .collect();

                    Some(OelTableRow { cells })
                })
                .collect();

            Some(OelBlock::Table(OelTable {
                id,
                rows,
                props: Default::default(),
            }))
        }
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Serde-based extractors for private fields in docx-rs types.
//
// docx-rs types with private fields implement Serialize and produce primitive
// JSON values (bool, number, string), NOT objects with a "val" key.
//   Bold    → serialize_bool(val)
//   Italic  → serialize_bool(val)
//   Sz      → serialize_u32(val)
//   Color   → serialize_str(val)
//   Highlight → serialize_str(val)
//   RunFonts → {"ascii": "...", "hiAnsi": "...", ...}  (camelCase, fields skipped if None)
//   LineSpacing → {"before": u32, "after": u32, ...}   (camelCase, fields skipped if None)
//   PageSize → {"w": u32, "h": u32}                    (private fields, camelCase)
// ---------------------------------------------------------------------------

fn serde_bool<T: serde::Serialize>(v: &T) -> bool {
    serde_json::to_value(v)
        .ok()
        .and_then(|j| j.as_bool())
        .unwrap_or(false)
}

fn serde_u64<T: serde::Serialize>(v: &T) -> Option<u64> {
    serde_json::to_value(v).ok().and_then(|j| j.as_u64())
}

fn serde_str<T: serde::Serialize>(v: &T) -> Option<String> {
    serde_json::to_value(v)
        .ok()
        .and_then(|j| j.as_str().map(|s| s.to_string()))
}

fn serde_str_field<T: serde::Serialize>(v: &T, key: &str) -> Option<String> {
    serde_json::to_value(v)
        .ok()
        .and_then(|j| j.get(key).and_then(|f| f.as_str()).map(|s| s.to_string()))
}

fn serde_u32_field<T: serde::Serialize>(v: &T, key: &str) -> Option<u32> {
    serde_json::to_value(v)
        .ok()
        .and_then(|j| j.get(key).and_then(|f| f.as_u64()).map(|n| n as u32))
}

fn convert_run_props(rp: &docx_rs::RunProperty) -> OelRunProps {
    OelRunProps {
        bold: rp.bold.as_ref().map(|b| serde_bool(b)).unwrap_or(false),
        italic: rp.italic.as_ref().map(|i| serde_bool(i)).unwrap_or(false),
        underline: rp.underline.is_some(),
        strikethrough: rp.strike.as_ref().map(|s| s.val).unwrap_or(false),
        font_size: rp.sz.as_ref().and_then(|s| serde_u64(s)).map(|v| v as u32),
        font_family: rp.fonts.as_ref().and_then(|f| serde_str_field(f, "ascii")),
        color: rp
            .color
            .as_ref()
            .and_then(|c| serde_str(c))
            .filter(|s| s != "auto" && !s.is_empty()),
        highlight: rp
            .highlight
            .as_ref()
            .and_then(|h| serde_str(h))
            .filter(|s| s != "none"),
    }
}

fn convert_para_props(pp: &docx_rs::ParagraphProperty) -> OelParaProps {
    // Justification.val is a public String ("left", "center", "right", "both", etc.)
    let alignment = pp
        .alignment
        .as_ref()
        .map(|j| match j.val.as_str() {
            "center" => Alignment::Center,
            "right" | "end" => Alignment::Right,
            "both" | "distribute" | "highKashida" | "lowKashida" | "mediumKashida" => {
                Alignment::Justify
            }
            _ => Alignment::Left,
        })
        .unwrap_or_default();

    let (list_type, indent_level) = if let Some(num) = &pp.numbering_property {
        let level = num.level.as_ref().map(|l| l.val as u32).unwrap_or(0);
        (Some(ListType::Bullet), level)
    } else {
        (None, 0)
    };

    // LineSpacing fields (before, after) are private; extract via serde (camelCase keys)
    let spacing_before = pp
        .line_spacing
        .as_ref()
        .and_then(|s| serde_u32_field(s, "before"));
    let spacing_after = pp
        .line_spacing
        .as_ref()
        .and_then(|s| serde_u32_field(s, "after"));

    OelParaProps {
        alignment,
        indent_level,
        list_type,
        spacing_before,
        spacing_after,
        line_spacing: None,
        style_id: pp.style.as_ref().map(|s| s.val.clone()),
    }
}

fn convert_section(docx: &Docx) -> OelSectionProps {
    let sp = &docx.document.section_property;

    // PageSize has private w/h fields; extract via serde (camelCase)
    let (page_width, page_height) = serde_json::to_value(&sp.page_size)
        .ok()
        .and_then(|j| {
            let w = j.get("w")?.as_u64()? as u32;
            let h = j.get("h")?.as_u64()? as u32;
            Some((w, h))
        })
        .unwrap_or((11906, 16838));

    // PageMargin fields are public i32
    let m = &sp.page_margin;
    OelSectionProps {
        page_width,
        page_height,
        margin_top: m.top.unsigned_abs(),
        margin_right: m.right.unsigned_abs(),
        margin_bottom: m.bottom.unsigned_abs(),
        margin_left: m.left.unsigned_abs(),
    }
}
