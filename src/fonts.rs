use std::collections::HashMap;
use std::io::{Cursor, Read};
use zip::ZipArchive;

/// Represents an extracted, de-obfuscated font from a DOCX file.
pub struct ExtractedFont {
    pub name: String,
    pub style: String,
    pub data: Vec<u8>,
}

/// Extracts and de-obfuscates embedded fonts from a DOCX byte slice.
pub fn extract_fonts(docx_bytes: &[u8]) -> Result<Vec<ExtractedFont>, String> {
    let cursor = Cursor::new(docx_bytes);
    let mut archive = ZipArchive::new(cursor).map_err(|e| format!("Failed to open zip: {}", e))?;

    // 1. Read fontTable.xml
    let mut font_table_xml = String::new();
    if let Ok(mut f) = archive.by_name("word/fontTable.xml") {
        f.read_to_string(&mut font_table_xml).map_err(|_| "Failed to read fontTable.xml")?;
    } else {
        return Ok(vec![]); // No embedded fonts
    }

    // 2. Read _rels/fontTable.xml.rels
    let mut font_table_rels = String::new();
    if let Ok(mut f) = archive.by_name("word/_rels/fontTable.xml.rels") {
        f.read_to_string(&mut font_table_rels).map_err(|_| "Failed to read fontTable.xml.rels")?;
    } else {
        return Ok(vec![]);
    }

    // 3. Map Relationship Id -> Target (e.g. "rId1" -> "fonts/font1.odttf")
    let mut rel_to_target = HashMap::new();
    // Super basic XML parsing for relationships
    for part in font_table_rels.split("<Relationship ") {
        if !part.contains("Id=") || !part.contains("Target=") { continue; }
        
        let id_start = part.find("Id=\"").unwrap() + 4;
        let id_end = part[id_start..].find("\"").unwrap() + id_start;
        let id = &part[id_start..id_end];

        let target_start = part.find("Target=\"").unwrap() + 8;
        let target_end = part[target_start..].find("\"").unwrap() + target_start;
        let target = &part[target_start..target_end];

        rel_to_target.insert(id.to_string(), target.to_string());
    }

    // 4. Map Target -> (Font Name, Style, FontKey)
    let mut target_to_info = HashMap::new();
    
    // Very basic parsing for fontTable.xml
    for font_part in font_table_xml.split("<w:font ") {
        if !font_part.contains("w:name=\"") { continue; }
        let name_start = font_part.find("w:name=\"").unwrap() + 8;
        let name_end = font_part[name_start..].find("\"").unwrap() + name_start;
        let font_name = &font_part[name_start..name_end];

        // find embeds
        let embeds = [
            ("Regular", "<w:embedRegular "),
            ("Bold", "<w:embedBold "),
            ("Italic", "<w:embedItalic "),
            ("BoldItalic", "<w:embedBoldItalic ")
        ];

        for (style, tag) in embeds.iter() {
            if let Some(idx) = font_part.find(tag) {
                let part = &font_part[idx..];
                if let (Some(r_idx), Some(k_idx)) = (part.find("r:id=\""), part.find("w:fontKey=\"")) {
                    let id_start = r_idx + 6;
                    let id_end = part[id_start..].find("\"").unwrap() + id_start;
                    let r_id = &part[id_start..id_end];

                    let key_start = k_idx + 11;
                    let key_end = part[key_start..].find("\"").unwrap() + key_start;
                    let font_key = &part[key_start..key_end];

                    if let Some(target) = rel_to_target.get(r_id) {
                        target_to_info.insert(target.to_string(), (font_name.to_string(), style.to_string(), font_key.to_string()));
                    }
                }
            }
        }
    }

    let mut extracted_fonts = Vec::new();

    // 5. Extract and De-obfuscate
    for (target, (name, style, font_key)) in target_to_info {
        // ZipArchive paths usually use forward slashes and might not include 'word/' if Target is relative.
        // Target is usually "fonts/font1.odttf", so we prepend "word/"
        let zip_path = format!("word/{}", target);
        
        let mut font_file = match archive.by_name(&zip_path) {
            Ok(f) => f,
            Err(_) => continue,
        };

        let mut data = Vec::new();
        font_file.read_to_end(&mut data).map_err(|_| "Failed to read font file")?;

        // De-obfuscate
        // Extract GUID from {UUID}
        let clean_key = font_key.replace("{", "").replace("}", "").replace("-", "");
        if clean_key.len() == 32 {
            let mut key_bytes = [0u8; 16];
            for i in 0..16 {
                key_bytes[i] = u8::from_str_radix(&clean_key[i*2..i*2+2], 16).unwrap_or(0);
            }
            
            // The XOR key is the GUID bytes reversed!
            // According to MS-ODRAWXML 4.2.1, the 16 bytes are reversed.
            let mut xor_key = [0u8; 32];
            for i in 0..16 {
                xor_key[i] = key_bytes[15 - i];
                xor_key[i + 16] = key_bytes[15 - i];
            }

            // XOR first 32 bytes
            for i in 0..std::cmp::min(32, data.len()) {
                data[i] ^= xor_key[i];
            }
        }

        extracted_fonts.push(ExtractedFont {
            name,
            style,
            data,
        });
    }

    Ok(extracted_fonts)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    #[test]
    fn test_extract_fonts() {
        let bytes = fs::read("tests/random/C53_L04A_Proprietes_Unite_dorganisation.docx").unwrap();
        let fonts = extract_fonts(&bytes).unwrap();
        
        assert!(!fonts.is_empty());
        for font in &fonts {
            println!("Extracted: {} ({}) - {} bytes", font.name, font.style, font.data.len());
            // Optionally, check if it starts with standard TTF/OTF magic bytes
            // TTF: 0x00, 0x01, 0x00, 0x00
            // OTF: 'O', 'T', 'T', 'O'
            let magic = &font.data[0..4];
            println!("  Magic: {:02x?} ({})", magic, String::from_utf8_lossy(magic));
            
            // To test, we can dump the first one to a file
            fs::write(format!("{}_{}.ttf", font.name, font.style), &font.data).unwrap();
        }
    }
}
