use crate::model::{
    Alignment, ListType, OelBlock, OelDocument, OelParaProps, OelParagraph, OelRun, OelRunProps,
    OelSectionProps, OelStyle, OelTable, OelTableCell, OelTableCellProps, OelTableRow, next_id,
};
use docx_rs::{
    DocumentChild, Docx, ParagraphChild, RunChild, TableCellContent, TableChild, TableRowChild,
};
use std::collections::HashMap;

pub fn docx_to_oel(docx: &Docx) -> OelDocument {
    // Build numbering lookup tables from word/numbering.xml so we can resolve
    // the correct ListType (Bullet vs Numbered) for each paragraph.
    //
    // DOCX numbering hierarchy:
    //   <w:num w:numId="N">          → references an abstractNum
    //     <w:abstractNumId w:val="M"/>
    //   <w:abstractNum w:abstractNumId="M">
    //     <w:lvl w:ilvl="0">
    //       <w:numFmt w:val="bullet"/>   ← "bullet" | "decimal" | "lowerLetter" | …
    //
    // We build: numId → abstractNumId, then (abstractNumId, level) → numFmt string.
    let mut num_id_to_abstract: HashMap<usize, usize> = HashMap::new();
    for num in &docx.numberings.numberings {
        num_id_to_abstract.insert(num.id, num.abstract_num_id);
    }
    let mut abstract_level_fmt: HashMap<(usize, usize), String> = HashMap::new();
    for abs_num in &docx.numberings.abstract_nums {
        for level in &abs_num.levels {
            abstract_level_fmt.insert((abs_num.id, level.level), level.format.val.clone());
        }
    }

    let blocks = docx
        .document
        .children
        .iter()
        .filter_map(|c| convert_document_child(c, &num_id_to_abstract, &abstract_level_fmt))
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
        // Styles don't reference numbering.xml instances, so an empty lookup is fine here.
        let para_props = convert_para_props(&style.paragraph_property, &HashMap::new(), &HashMap::new());

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

fn convert_document_child(
    child: &DocumentChild,
    num_id_to_abstract: &HashMap<usize, usize>,
    abstract_level_fmt: &HashMap<(usize, usize), String>,
) -> Option<OelBlock> {
    match child {
        DocumentChild::Paragraph(p) => {
            let mut para = OelParagraph::new(next_id());
            para.props = convert_para_props(&p.property, num_id_to_abstract, abstract_level_fmt);

            for pc in &p.children {
                if let ParagraphChild::Run(run) = pc {
                    let props = convert_run_props(&run.run_property);
                    let mut text = String::new();
                    for rc in &run.children {
                        match rc {
                            RunChild::Text(t) => text.push_str(&t.text),
                            RunChild::Tab(_) => text.push('\t'),
                            RunChild::Break(_) => text.push('\n'),
                            RunChild::Drawing(d) => {
                                if let Some(docx_rs::DrawingData::Pic(pic)) = &d.data {
                                    let drawing = convert_drawing(pic);
                                    if !text.is_empty() {
                                        para.runs.push(OelRun::with_props(
                                            std::mem::take(&mut text),
                                            props.clone(),
                                        ));
                                    }
                                    para.runs.push(OelRun::with_drawing(drawing, props.clone()));
                                }
                            }
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
                                        para.props = convert_para_props(&p.property, num_id_to_abstract, abstract_level_fmt);
                                        for pc in &p.children {
                                            if let ParagraphChild::Run(run) = pc {
                                                let props = convert_run_props(&run.run_property);
                                                let mut text = String::new();
                                                for rc in &run.children {
                                                    match rc {
                                                        RunChild::Text(t) => text.push_str(&t.text),
                                                        RunChild::Tab(_) => text.push('\t'),
                                                        RunChild::Break(_) => text.push('\n'),
                                                        RunChild::Drawing(d) => {
                                                            if let Some(docx_rs::DrawingData::Pic(pic)) = &d.data {
                                                                let drawing = convert_drawing(pic);
                                                                if !text.is_empty() {
                                                                    para.runs.push(OelRun::with_props(std::mem::take(&mut text), props.clone()));
                                                                }
                                                                para.runs.push(OelRun::with_drawing(drawing, props.clone()));
                                                            }
                                                        }
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

fn convert_para_props(
    pp: &docx_rs::ParagraphProperty,
    num_id_to_abstract: &HashMap<usize, usize>,
    abstract_level_fmt: &HashMap<(usize, usize), String>,
) -> OelParaProps {
    // Justification.val is a public String ("left", "center", "right", "both", etc.).
    // Keep None when absent so style-level alignment can be inherited in the renderer.
    let alignment = pp.alignment.as_ref().map(|j| match j.val.as_str() {
        "center" => Alignment::Center,
        "right" | "end" => Alignment::Right,
        "both" | "distribute" | "highKashida" | "lowKashida" | "mediumKashida" => {
            Alignment::Justify
        }
        _ => Alignment::Left,
    });

    // Resolve ListType from numbering.xml instead of assuming everything is Bullet.
    // DOCX: numFmt="bullet" → Bullet; any other format (decimal, lowerLetter…) → Numbered.
    let (list_type, indent_level, num_id) = if let Some(num) = &pp.numbering_property {
        let raw_num_id = num.id.as_ref().map(|n| n.id).unwrap_or(0);
        let level = num.level.as_ref().map(|l| l.val as u32).unwrap_or(0);

        let num_fmt = num_id_to_abstract
            .get(&raw_num_id)
            .and_then(|abs_id| abstract_level_fmt.get(&(*abs_id, level as usize)))
            .map(|s| s.as_str());

        let list_type = match num_fmt {
            Some("bullet") => ListType::Bullet,
            _ => ListType::Numbered,
        };

        (Some(list_type), level, Some(raw_num_id as u32))
    } else {
        (None, 0, None)
    };

    // LineSpacing fields are private; extract via serde (camelCase keys).
    // `line` is in 240ths-of-a-line when lineRule="auto" (240=1×, 360=1.5×, 480=2×).
    // `before`/`after` are in twips.
    let spacing_before = pp
        .line_spacing
        .as_ref()
        .and_then(|s| serde_u32_field(s, "before"));
    let spacing_after = pp
        .line_spacing
        .as_ref()
        .and_then(|s| serde_u32_field(s, "after"));
    let line_spacing = pp.line_spacing.as_ref().and_then(|s| {
        let json = serde_json::to_value(s).ok()?;
        let line = json.get("line")?.as_i64()? as f32;
        let rule = json
            .get("lineRule")
            .and_then(|r| r.as_str())
            .unwrap_or("auto");
        // Only "auto" maps cleanly to a unitless CSS multiplier.
        // "exact" / "atLeast" are absolute twip heights — skip for now.
        if rule == "auto" { Some(line / 240.0) } else { None }
    });

    OelParaProps {
        alignment,
        indent_level,
        list_type,
        num_id,
        spacing_before,
        spacing_after,
        line_spacing,
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

fn convert_drawing(pic: &docx_rs::Pic) -> crate::model::block::OelDrawing {
    let is_floating = matches!(pic.position_type, docx_rs::DrawingPositionType::Anchor);
    let width_pt = pic.size.0 as f32 / 12700.0;
    let height_pt = pic.size.1 as f32 / 12700.0;
    let (offset_x_pt, offset_y_pt) = if is_floating {
        let ox = match &pic.position_h {
            docx_rs::DrawingPosition::Offset(x) => *x as f32 / 12700.0,
            _ => 0.0,
        };
        let oy = match &pic.position_v {
            docx_rs::DrawingPosition::Offset(y) => *y as f32 / 12700.0,
            _ => 0.0,
        };
        (ox, oy)
    } else {
        (0.0, 0.0)
    };
    let relative_from_h = pic.relative_from_h.to_string();
    let relative_from_v = pic.relative_from_v.to_string();
    crate::model::block::OelDrawing {
        id: pic.id.clone(),
        width_pt,
        height_pt,
        is_floating,
        offset_x_pt,
        offset_y_pt,
        wrapping_mode: convert_wrapping_mode(pic),
        relative_from_h,
        relative_from_v,
        z_order: pic.relative_height,
    }
}

fn convert_wrapping_mode(pic: &docx_rs::Pic) -> crate::model::block::OelWrappingMode {
    use crate::model::block::OelWrappingMode;
    
    // Pic.position_type is public: Anchor or Inline
    if matches!(pic.position_type, docx_rs::DrawingPositionType::Inline) {
        return OelWrappingMode::Inline;
    }

    // pic.wrap is private in docx-rs 0.4. Extract it via serde.
    if let Ok(json) = serde_json::to_value(pic) {
        if let Some(wrap) = json.get("wrap") {
            if let Some(mode) = wrap.as_str() {
                return match mode {
                    "square" => OelWrappingMode::Square,
                    "tight" => OelWrappingMode::Tight,
                    "through" => OelWrappingMode::Through,
                    "topAndBottom" => OelWrappingMode::TopAndBottom,
                    "behindText" => OelWrappingMode::BehindText,
                    "inFrontOfText" => OelWrappingMode::InFrontOfText,
                    _ => OelWrappingMode::Square, // Default for anchored
                };
            }
        }
    }
    
    OelWrappingMode::Square // Default fallback for floating
}
