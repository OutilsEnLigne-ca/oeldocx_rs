fn main() {
    use std::fs;
    use std::io::Read;
    use zip::ZipArchive;

    let file = fs::File::open("tests/random/C53_L07A_Proprietes_Groupe.docx").unwrap();
    let mut archive = ZipArchive::new(file).unwrap();
    
    let mut font_table_xml = String::new();
    if let Ok(mut f) = archive.by_name("word/fontTable.xml") {
        f.read_to_string(&mut font_table_xml).unwrap();
    }
    
    let mut font_table_rels = String::new();
    if let Ok(mut f) = archive.by_name("word/_rels/fontTable.xml.rels") {
        f.read_to_string(&mut font_table_rels).unwrap();
    }
    
    println!("fontTable.xml:\n{}", font_table_xml);
    println!("\nfontTable.xml.rels:\n{}", font_table_rels);
}
#[test]
fn test_doc_defaults() {
    let docx = docx_rs::Docx::new();
    let defaults = docx.styles.doc_defaults;
}
#[test]
fn test_doc_defaults_print() {
    let docx = docx_rs::Docx::new();
    println!("{:#?}", docx.styles.doc_defaults);
}
#[test]
fn test_drawing_struct() {
    let d = docx_rs::Drawing::new();
}
#[test]
fn test_drawing_print() {
    let d = docx_rs::Drawing::new();
    println!("{:#?}", d);
}
#[test]
fn test_pic_print() {
    let p = docx_rs::Pic::new_with_dimensions(vec![], 0, 0);
    println!("{:#?}", p);
}
#[test]
fn test_drawing_data_enum() {
    let d: Option<docx_rs::DrawingData> = None;
}

#[test]
fn test_image_extraction() {
    let bytes = std::fs::read("tests/random/C61_-_Document_de_definition.docx").unwrap();
    let docx = docx_rs::read_docx(&bytes).unwrap();
    println!("Images count: {}", docx.images.len());
    for (id, path, img, png) in docx.images {
        println!("Image: {} -> {} | img size: {}, png size: {}", id, path, img.0.len(), png.0.len());
    }
}
#[test]
fn test_image_extraction_embed() {
    let bytes = std::fs::read("tests/Images/embed_images.docx").unwrap();
    let mut docx = docx_rs::read_docx(&bytes).unwrap();
    println!("Images count: {}", docx.images.len());
    for (id, path, img, png) in docx.images {
        println!("Image: {} -> {} | img size: {}, png size: {}", id, path, img.0.len(), png.0.len());
    }
}
#[test]
fn test_image_size() {
    let bytes = std::fs::read("tests/Images/embed_images.docx").unwrap();
    let mut docx = docx_rs::read_docx(&bytes).unwrap();
    
    use docx_rs::{DocumentChild, ParagraphChild, RunChild};
    for child in &docx.document.children {
        if let DocumentChild::Paragraph(p) = child {
            for pc in &p.children {
                if let ParagraphChild::Run(run) = pc {
                    for rc in &run.children {
                        if let RunChild::Drawing(d) = rc {
                            println!("{:#?}", d.data);
                        }
                    }
                }
            }
        }
    }
}
