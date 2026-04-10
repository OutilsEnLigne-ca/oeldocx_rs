use crate::model::{ListType, OelBlock, OelDocument, OelParagraph, OelRunProps, OelTable};
use crate::render::{
    DEFAULT_COLOR, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE_PT, HALFPT_TO_PT, RenderBlock,
    RenderDocument, RenderDrawing, RenderFormat, RenderParagraph, RenderSectionProps, RenderSpan,
    RenderTable, RenderTableCell, RenderTableRow, TWIPS_TO_PT,
};

pub fn oel_to_render(doc: &OelDocument) -> RenderDocument {
    let section = RenderSectionProps {
        page_width_pt: doc.section.page_width as f32 * TWIPS_TO_PT,
        page_height_pt: doc.section.page_height as f32 * TWIPS_TO_PT,
        margin_top_pt: doc.section.margin_top as f32 * TWIPS_TO_PT,
        margin_right_pt: doc.section.margin_right as f32 * TWIPS_TO_PT,
        margin_bottom_pt: doc.section.margin_bottom as f32 * TWIPS_TO_PT,
        margin_left_pt: doc.section.margin_left as f32 * TWIPS_TO_PT,
    };

    let mut numbered_counter: u32 = 0;

    let blocks = doc
        .blocks
        .iter()
        .map(|block| convert_block(block, doc, &mut numbered_counter))
        .collect();

    RenderDocument { blocks, section }
}

fn convert_block(block: &OelBlock, doc: &OelDocument, numbered_counter: &mut u32) -> RenderBlock {
    match block {
        OelBlock::Paragraph(p) => {
            RenderBlock::Paragraph(convert_paragraph(p, doc, numbered_counter))
        }
        OelBlock::Table(t) => RenderBlock::Table(convert_table(t, doc, numbered_counter)),
        OelBlock::PageBreak => RenderBlock::PageBreak,
    }
}

fn convert_paragraph(
    para: &OelParagraph,
    doc: &OelDocument,
    numbered_counter: &mut u32,
) -> RenderParagraph {
    let list_index = match &para.props.list_type {
        Some(ListType::Numbered) => {
            *numbered_counter += 1;
            Some(*numbered_counter)
        }
        Some(ListType::Bullet) => {
            *numbered_counter = 0;
            None
        }
        None => {
            *numbered_counter = 0;
            None
        }
    };

    let spacing_before_pt = para
        .props
        .spacing_before
        .map(|v| v as f32 * TWIPS_TO_PT)
        .unwrap_or(0.0);
    let spacing_after_pt = para
        .props
        .spacing_after
        .map(|v| v as f32 * TWIPS_TO_PT)
        .unwrap_or(0.0);

    let mut char_cursor: usize = 0;
    let spans: Vec<RenderSpan> = para
        .runs
        .iter()
        .map(|run| {
            let char_start = char_cursor;
            let len = run.char_len();
            let char_end = char_start + len;
            char_cursor = char_end;

            RenderSpan {
                text: run.text.clone(),
                drawing: run.drawing.as_ref().map(|d| RenderDrawing {
                    id: d.id.clone(),
                    width_pt: d.width_pt,
                    height_pt: d.height_pt,
                    is_floating: d.is_floating,
                    offset_x_pt: d.offset_x_pt,
                    offset_y_pt: d.offset_y_pt,
                    wrapping_mode: match d.wrapping_mode {
                        crate::model::block::OelWrappingMode::Inline => {
                            crate::render::RenderWrappingMode::Inline
                        }
                        crate::model::block::OelWrappingMode::Square => {
                            crate::render::RenderWrappingMode::Square
                        }
                        crate::model::block::OelWrappingMode::Tight => {
                            crate::render::RenderWrappingMode::Tight
                        }
                        crate::model::block::OelWrappingMode::Through => {
                            crate::render::RenderWrappingMode::Through
                        }
                        crate::model::block::OelWrappingMode::TopAndBottom => {
                            crate::render::RenderWrappingMode::TopAndBottom
                        }
                        crate::model::block::OelWrappingMode::BehindText => {
                            crate::render::RenderWrappingMode::BehindText
                        }
                        crate::model::block::OelWrappingMode::InFrontOfText => {
                            crate::render::RenderWrappingMode::InFrontOfText
                        }
                    },
                }),
                format: resolve_format(&run.props, para.props.style_id.as_deref(), doc),
                char_start,
                char_end,
            }
        })
        .collect();

    RenderParagraph {
        id: para.id.clone(),
        style_id: para.props.style_id.clone(),
        alignment: para.props.alignment.clone(),
        indent_level: para.props.indent_level,
        list_type: para.props.list_type.clone(),
        list_index,
        spacing_before_pt,
        spacing_after_pt,
        line_spacing: para.props.line_spacing,
        spans,
    }
}

fn convert_table(table: &OelTable, doc: &OelDocument, numbered_counter: &mut u32) -> RenderTable {
    let rows = table
        .rows
        .iter()
        .map(|row| {
            let cells = row
                .cells
                .iter()
                .map(|cell| {
                    let blocks = cell
                        .blocks
                        .iter()
                        .map(|b| convert_block(b, doc, numbered_counter))
                        .collect();
                    RenderTableCell {
                        blocks,
                        col_span: cell.props.col_span,
                        row_span: cell.props.row_span,
                    }
                })
                .collect();
            RenderTableRow { cells }
        })
        .collect();

    RenderTable {
        id: table.id.clone(),
        rows,
    }
}

fn resolve_format(props: &OelRunProps, style_id: Option<&str>, doc: &OelDocument) -> RenderFormat {
    let style_run_props = style_id
        .and_then(|id| doc.styles.get(id))
        .map(|s| &s.run_props);

    let bold = props.bold || style_run_props.map_or(false, |sp| sp.bold);
    let italic = props.italic || style_run_props.map_or(false, |sp| sp.italic);
    let underline = props.underline || style_run_props.map_or(false, |sp| sp.underline);
    let strikethrough = props.strikethrough || style_run_props.map_or(false, |sp| sp.strikethrough);

    let font_size = props
        .font_size
        .or_else(|| style_run_props.and_then(|sp| sp.font_size));
    let font_family = props
        .font_family
        .as_ref()
        .or_else(|| style_run_props.and_then(|sp| sp.font_family.as_ref()));
    let color = props
        .color
        .as_ref()
        .or_else(|| style_run_props.and_then(|sp| sp.color.as_ref()));

    RenderFormat {
        bold,
        italic,
        underline,
        strikethrough,
        font_size_pt: font_size
            .map(|s| s as f32 * HALFPT_TO_PT)
            .unwrap_or(DEFAULT_FONT_SIZE_PT),
        font_family: font_family
            .cloned()
            .unwrap_or_else(|| DEFAULT_FONT_FAMILY.to_string()),
        color: color.cloned().unwrap_or_else(|| DEFAULT_COLOR.to_string()),
        highlight: props.highlight.clone(),
    }
}
