use zip::read::{ZipArchive, ZipFile};
use zip::result::ZipError;
use std::fs::File;
use std::path::Path;
use std::result::Result;
use std::error::Error;
use std::io::{Read, Cursor, Result as IOResult};

pub struct Archive(ZipArchive<Cursor<Vec<u8>>>);

impl Archive {
    pub fn new<P: AsRef<Path>>(path: P) -> Result<Archive, Box<dyn Error>> {
        let data = File::open(path)?.read_all_to_vec()?;
        Ok(Archive(ZipArchive::new(Cursor::new(data))?))
    }

    pub fn by_name(&mut self, name: &str) -> Result<ZipFile, ZipError> {
        self.0.by_name(name)
    }

    pub fn file_names<'a>(&self) -> impl Iterator<Item=&str> {
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
fn test_archive() -> Result<(), Box<dyn Error>> {
    let mut ar = Archive::new(super::test_file()).unwrap();
    for name in ar.file_names() {
        println!("{}", name);
    }
    let mut file = ar.by_name("xl/sharedStrings.xml")?;
    println!("{}", file.read_all_to_string()?);
    Ok(())
}