//extern crate zip;

use std::fs;
use std::io::{Read, Seek};
use std::collections::BTreeMap as Map;

use xml;
use refer;
use xlsx;
use result::{XlsxResult, Error};
use zip;

#[derive(Debug)]
pub struct File {
    rels: Map<String, String>,
    strs: refer::Strings,
    clrs: refer::Colors,
    nfts: refer::NumFmts,

    xml_styles: Option<xml::styles::StyleSheet>,

    workbook: xlsx::WorkBook,
}

impl File {
    fn open(file_name: &str) -> XlsxResult<Self> {
        // step 1: open file
        let file = fs::File::open(file_name)?;

        // step 2: open zip
        let mut zip = zip::ZipArchive::new(file)?;

        // step 3: process xmls
        let mut xlsx_file = File {
            rels: Map::new(),
            strs: refer::Strings::new(),
            clrs: refer::Colors::new(),
            nfts: refer::NumFmts::new(),

            xml_styles: None,

            workbook: xlsx::WorkBook::new(),
        };

        for i in 0..zip.len() {
            let mut f = zip.by_index(i).unwrap();
            println!("Filename: {}", f.name());
            match f.name() {
                "xl/_rels/workbook.xml.rels" => xlsx_file.load_rels(f)?,
                "xl/sharedStrings.xml" => xlsx_file.load_strs(f)?,
                "xl/theme/theme1.xml" => xlsx_file.load_theme(f)?,
                "xl/styles.xml" => xlsx_file.load_style(f)?,
                _ => (),
            }
        }
        let xml_wb = xml::workbook::Workbook::from_xml(zip.by_name("xl/workbook.xml")?)?;

        xlsx_file.workbook.date1904 = xml_wb.workbookPr.date1904.parse().unwrap_or(false);

        for sheet in xml_wb.sheets.items {
            xlsx_file.load_sheet(&mut zip, &sheet)?;
        }


        Ok(xlsx_file)
    }

    fn load_rels<R: Read>(&mut self, reader: R) -> XlsxResult<()> {
        match xml::rels::Relationships::from_xml(reader) {
            Ok(rels) => {
                for r in rels.items {
                    self.rels.insert(r.id.clone(), r.target.clone());
                }
                Ok(())
            }
            Err(err) => Err(Error::Xlsx(format!("load rels error: {}", err))),
        }
    }

    fn load_strs<R: Read>(&mut self, reader: R) -> XlsxResult<()> {
        match xml::shared_strings::SharedStrings::from_xml(reader) {
            Ok(sst) => {
                for si in sst.items {
                    self.strs.add(&si.t);
                }
                Ok(())
            }
            Err(err) => Err(Error::Xlsx(format!("load shared_strings error: {}", err))),
        }
    }

    fn load_theme<R: Read>(&mut self, reader: R) -> XlsxResult<()> {
        match xml::theme::Theme::from_xml(reader) {
            Ok(thm) => {
                let ct = thm.themeElements.clrScheme;
                for (name, clr) in ct {
                    self.clrs.insert(name, clr.rgb_color());
                }
                Ok(())
            }
            Err(err) => Err(Error::Xlsx(format!("load theme error: {}", err))),
        }
    }

    fn load_style<R: Read>(&mut self, reader: R) -> XlsxResult<()> {
        match xml::styles::StyleSheet::from_xml(reader) {
            Ok(ss) => {
                match ss.numFmts {
                    Some(ref nfs) => {
                        for nf in &nfs.items {
                            self.nfts.insert(&nf.numFmtId, &nf.formatCode);
                        }
                    }
                    _ => (),
                }
                self.xml_styles = Some(ss);
                Ok(())
            }
            Err(err) => Err(Error::Xlsx(format!("load style error: {}", err))),
        }
    }

    fn load_sheet<R: Read + Seek>(
        &mut self,
        zip: &mut zip::ZipArchive<R>,
        sheet: &xml::workbook::Sheet,
    ) -> XlsxResult<()> {
        let sheet_file = match self.rels.get(&sheet.id) {
            Some(s) => format!("xl/{}", s),
            None => {
                if sheet.sheetId != "" {
                    format!("xl/worksheets/sheet{}", sheet.sheetId)
                } else {
                    format!("xl/worksheets/sheet{}", sheet.id)
                }
            }
        };
        //println!("load_sheet: {}", sheet_file);
        let xml_sheet = xml::sheet::Worksheet::from_xml(zip.by_name(&sheet_file)?)?;
        //println!("{:#?}", xml_sheet);
        let xlsx_sheet = xlsx::Sheet::from_xml(xml_sheet, &self)?;
        self.workbook.insert(&sheet.name, xlsx_sheet);
        Ok(())
    }

    pub fn get_num_fmt(&self, style_id: usize) -> Option<String> {
        if self.xml_styles.is_none() {
            return None;
        } else {
            let xf = self.xml_styles.as_ref().unwrap().cellXfs.items.get(style_id);
            if xf.is_none() {
                return None;
            } else {
                match self.nfts.get(&xf.unwrap().numFmtId) {
                    Some(s) => Some(s.clone()),
                    None => None,
                }
            }
        }
    }
}

#[test]
fn test_file_open() {
    let _f = File::open(&format!("{}/tests/table.xlsx", env!("CARGO_MANIFEST_DIR")));
    println!("{:#?}", _f.unwrap().workbook);
}
