use std::io::{Cursor, Write, Seek};
use docx_rs::{
    self, AlignmentType, AbstractNumbering, Level, NumberFormat, LevelText, LevelJc,
    Numbering, NumberingId, IndentLevel, Start,
};
use crate::model::{
    OelDocument, OelBlock, OelParagraph, OelRun, OelTable, OelTableCell,
    OelRunProps, OelParaProps, Alignment, ListType,
};

const BULLET_ABSTRACT_NUM_ID: usize = 1;
const NUMBERED_ABSTRACT_NUM_ID: usize = 2;
const BULLET_NUM_ID: usize = 1;
const NUMBERED_NUM_ID: usize = 2;

#[derive(Debug)]
pub struct SerializeError(pub String);

impl std::fmt::Display for SerializeError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "serialize error: {}", self.0)
    }
}

pub fn oel_to_bytes(doc: &OelDocument) -> Result<Vec<u8>, SerializeError> {
    let mut docx = build_docx(doc);

    let mut buf = Cursor::new(Vec::new());
    docx.build()
        .pack(&mut buf)
        .map_err(|e| SerializeError(e.to_string()))?;
    Ok(buf.into_inner())
}

fn build_docx(doc: &OelDocument) -> docx_rs::Docx {
    let has_bullets = has_list_type(doc, &ListType::Bullet);
    let has_numbered = has_list_type(doc, &ListType::Numbered);

    let mut d = docx_rs::Docx::new();

    if has_bullets {
        d = d
            .add_abstract_numbering(
                AbstractNumbering::new(BULLET_ABSTRACT_NUM_ID)
                    .add_level(
                        Level::new(
                            0,
                            Start::new(1),
                            NumberFormat::new("bullet"),
                            LevelText::new("•"),
                            LevelJc::new("left"),
                        )
                        .indent(Some(720), Some(docx_rs::SpecialIndentType::Hanging(360)), None, None),
                    )
                    .add_level(
                        Level::new(
                            1,
                            Start::new(1),
                            NumberFormat::new("bullet"),
                            LevelText::new("◦"),
                            LevelJc::new("left"),
                        )
                        .indent(Some(1440), Some(docx_rs::SpecialIndentType::Hanging(360)), None, None),
                    ),
            )
            .add_numbering(Numbering::new(BULLET_NUM_ID, BULLET_ABSTRACT_NUM_ID));
    }

    if has_numbered {
        d = d
            .add_abstract_numbering(
                AbstractNumbering::new(NUMBERED_ABSTRACT_NUM_ID)
                    .add_level(
                        Level::new(
                            0,
                            Start::new(1),
                            NumberFormat::new("decimal"),
                            LevelText::new("%1."),
                            LevelJc::new("left"),
                        )
                        .indent(Some(720), Some(docx_rs::SpecialIndentType::Hanging(360)), None, None),
                    )
                    .add_level(
                        Level::new(
                            1,
                            Start::new(1),
                            NumberFormat::new("lowerLetter"),
                            LevelText::new("%2."),
                            LevelJc::new("left"),
                        )
                        .indent(Some(1440), Some(docx_rs::SpecialIndentType::Hanging(360)), None, None),
                    ),
            )
            .add_numbering(Numbering::new(NUMBERED_NUM_ID, NUMBERED_ABSTRACT_NUM_ID));
    }

    for block in &doc.blocks {
        d = add_block(d, block);
    }

    d
}

fn add_block(mut d: docx_rs::Docx, block: &OelBlock) -> docx_rs::Docx {
    match block {
        OelBlock::Paragraph(p) => d.add_paragraph(build_paragraph(p)),
        OelBlock::Table(t) => d.add_table(build_table(t)),
        OelBlock::PageBreak => {
            let run = docx_rs::Run::new().add_break(docx_rs::BreakType::Page);
            d.add_paragraph(docx_rs::Paragraph::new().add_run(run))
        }
    }
}

fn build_paragraph(para: &OelParagraph) -> docx_rs::Paragraph {
    let mut p = docx_rs::Paragraph::new();

    for run in &para.runs {
        p = p.add_run(build_run(run));
    }

    p = apply_para_props(p, &para.props);
    p
}

fn apply_para_props(mut p: docx_rs::Paragraph, props: &OelParaProps) -> docx_rs::Paragraph {
    let align = match props.alignment {
        Alignment::Left => AlignmentType::Left,
        Alignment::Center => AlignmentType::Center,
        Alignment::Right => AlignmentType::Right,
        Alignment::Justify => AlignmentType::Both,
    };
    p = p.align(align);

    if let Some(list_type) = &props.list_type {
        let num_id = match list_type {
            ListType::Bullet => BULLET_NUM_ID,
            ListType::Numbered => NUMBERED_NUM_ID,
        };
        p = p.numbering(NumberingId::new(num_id), IndentLevel::new(props.indent_level as usize));
    }

    if let Some(style) = &props.style_id {
        p = p.style(style);
    }

    p
}

fn build_run(run: &OelRun) -> docx_rs::Run {
    let mut r = docx_rs::Run::new().add_text(&run.text);
    r = apply_run_props(r, &run.props);
    r
}

fn apply_run_props(mut r: docx_rs::Run, props: &OelRunProps) -> docx_rs::Run {
    if props.bold { r = r.bold(); }
    if props.italic { r = r.italic(); }
    if props.underline { r = r.underline("single"); }
    if props.strikethrough { r = r.strike(); }
    if let Some(sz) = props.font_size {
        r = r.size(sz as usize);
    }
    if let Some(color) = &props.color {
        r = r.color(color);
    }
    if let Some(family) = &props.font_family {
        r = r.fonts(docx_rs::RunFonts::new().ascii(family));
    }
    r
}

fn build_table(table: &OelTable) -> docx_rs::Table {
    let rows: Vec<docx_rs::TableRow> = table.rows.iter().map(|row| {
        let cells: Vec<docx_rs::TableCell> = row.cells.iter().map(build_table_cell).collect();
        docx_rs::TableRow::new(cells)
    }).collect();
    docx_rs::Table::new(rows)
}

fn build_table_cell(cell: &OelTableCell) -> docx_rs::TableCell {
    let mut tc = docx_rs::TableCell::new();
    for block in &cell.blocks {
        match block {
            OelBlock::Paragraph(p) => {
                tc = tc.add_paragraph(build_paragraph(p));
            }
            // Nested tables and page breaks inside cells are not supported in Phase 1.
            _ => {}
        }
    }
    tc
}

fn has_list_type(doc: &OelDocument, target: &ListType) -> bool {
    doc.blocks.iter().any(|b| match b {
        OelBlock::Paragraph(p) => p.props.list_type.as_ref() == Some(target),
        _ => false,
    })
}
