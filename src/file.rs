extern crate zip;

use std::fs;
use std::io::{Read};

use std::collections::BTreeMap as Map;

use xlsx;
        
#[derive(Debug)]
pub struct File {
    rels: Map<String, String>,
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
        };

        for i in 0..zip.len() {
            let mut f = zip.by_index(i).unwrap();
            println!("Filename: {}", f.name());
            match f.name() {
                "xl/_rels/workbook.xml.rels" => {
                    xlsx_file.load_rels(f)
                },
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

}

#[test]
fn test_file_open() {
    let f = File::open(&format!("{}/tests/table.xlsx", env!("CARGO_MANIFEST_DIR")));
    println!("{:?}", f);
}
