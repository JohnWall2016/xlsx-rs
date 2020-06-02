use zip::read::{ZipArchive, ZipFile};
use std::fs::File;
use std::path::Path;
use std::io::{Read, Cursor, Result as IOResult};

use super::XlsXResult;

pub struct Archive(ZipArchive<Cursor<Vec<u8>>>);

impl Archive {
    pub fn new<P: AsRef<Path>>(path: P) -> XlsXResult<Self> {
        let data = File::open(path)?.read_all_to_vec()?;
        Ok(Archive(ZipArchive::new(Cursor::new(data))?))
    }

    pub fn by_name<'a>(&'a mut self, name: &str) -> XlsXResult<ZipFile> {
        Ok(self.0.by_name(name)?)
    }

    pub fn file_names(&self) -> impl Iterator<Item=&str> {
        self.0.file_names()
    }
}

pub trait ReadAll {
    fn read_all_to_string(&mut self) -> IOResult<String>;
    fn read_all_to_vec(&mut self) -> IOResult<Vec<u8>>;
}

impl<T: Read> ReadAll for T {
    fn read_all_to_string(&mut self) -> IOResult<String> {
        let mut str = String::new();
        self.read_to_string(&mut str)?;
        Ok(str)
    }

    fn read_all_to_vec(&mut self) -> IOResult<Vec<u8>> {
        let mut buf = Vec::new();
        self.read_to_end(&mut buf)?;
        Ok(buf)
    }
}

#[test]
fn test_archive() -> XlsXResult<()> {
    let mut ar = Archive::new(super::test_file()).unwrap();
    for name in ar.file_names() {
        println!("{}", name);
    }
    {
        let mut file = ar.by_name("xl/sharedStrings.xml")?;
        println!("{}", file.read_all_to_string()?);
    }
    {
        let mut file = ar.by_name("xl/sharedStrings.xml")?;
        println!("{}", file.read_all_to_string()?);
    }
    Ok(())
}