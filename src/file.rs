extern crate zip;

use std::fs;
use std::io::{Read};

use std::collections::BTreeMap as Map;

use xlsx;

use refer;
        
#[derive(Debug)]
pub struct File {
    rels: Map<String, String>,
    strs: refer::Strings,
    clrs: refer::Colors,
}

impl File {
    fn open(file_name: &str) -> Self {
        // step 1: open file
        let file = match fs::File::open(file_name) {
            Ok(f) => f,
            Err(err) => panic!("open file error: {} {}", err, file_name),
        };

        // step 2: open zip
        let mut zip = match zip::ZipArchive::new(file) {
            Ok(z) => z,
            Err(err) => panic!("read zip error: {}", err),
        };

        // step 3: process xmls
        let mut xlsx_file = File {
            rels: Map::new(),
            strs: refer::Strings::new(),
            clrs: refer::Colors::new(),
        };

        for i in 0..zip.len() {
            let mut f = zip.by_index(i).unwrap();
            println!("Filename: {}", f.name());
            match f.name() {
                "xl/_rels/workbook.xml.rels" => {
                    xlsx_file.load_rels(f)
                },
                "xl/sharedStrings.xml" => {
                    xlsx_file.load_strs(f)
                },
                "xl/theme/theme1.xml" => {
                    xlsx_file.load_theme(f)
                }
                _ => (),
            }
        }

        xlsx_file
    }

    fn load_rels<R: Read>(self: &mut Self, reader: R) {
        match xlsx::rels::Relationships::from_xml(reader) {
            Ok(rels) => {
                for r in rels.items() {
                    self.rels.insert(r.id.clone(), r.target.clone());
                }
            },
            Err(err) => panic!("load rels error: {}", err),
        }
    }

    fn load_strs<R: Read>(self: &mut Self, reader: R) {
        match xlsx::shared_strings::SharedStrings::from_xml(reader) {
            Ok(sst) => {
                for si in sst.items() {
                    self.strs.add(&si.t);
                }
            },
            Err(err) => panic!("load shared_strings error: {}", err),
        }
    }

    fn load_theme<R: Read>(self: &mut Self, reader: R) {
        match xlsx::theme::Theme::from_xml(reader) {
            Ok(thm) => {
                let ct = thm.themeElements.clrScheme;
                for (name, clr) in ct {
                    self.clrs.insert(String::from(name), clr.rgb_color().clone());
                }
            },
            Err(err) => panic!("load theme error: {}", err),
        }
    }

}

#[test]
fn test_file_open() {
    let f = File::open(&format!("{}/tests/table.xlsx", env!("CARGO_MANIFEST_DIR")));
    println!("{:#?}", f);
}
