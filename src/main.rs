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
