
use std::io;
use zip;

use serde_xml_rs;

#[derive(Debug)]
pub enum Error {
    Io(io::Error),
    Zip(zip::result::ZipError),
    Xml(serde_xml_rs::Error),
    Xlsx(String),
}

impl Error {
    pub fn xlsx<T>(msg: &str) -> Result<T, Self> {
        Err(Error::Xlsx(msg.to_string()))
    }
}

use std::convert;

impl convert::From<io::Error> for Error {
    fn from(err: io::Error) -> Error {
        Error::Io(err)
    }
}

impl convert::From<zip::result::ZipError> for Error {
    fn from(err: zip::result::ZipError) -> Error {
        Error::Zip(err)
    }
}

impl convert::From<serde_xml_rs::Error> for Error {
    fn from(err: serde_xml_rs::Error) -> Error {
        Error::Xml(err)
    }
}

use std::fmt;
use std::error;

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter) -> Result<(), fmt::Error> {
        match *self {
            Error::Io(ref err) => {
                write!(f, "Io error: {}", (err as &error::Error).description())
            },
            Error::Zip(ref err) => {
                write!(f, "Zip error: {}", (err as &error::Error).description())
            },
            Error::Xml(ref err) => {
                write!(f, "Xml error: {}", (err as &error::Error).description())
            },
            Error::Xlsx(ref msg) => {
                write!(f, "Xlsx error: {}", msg)
            },
        }
    }
}

impl error::Error for Error {
    fn description(&self) -> &str {
        match *self {
            Error::Io(ref err) => (err as &error::Error).description(),
            Error::Zip(ref err) => (err as &error::Error).description(),
            Error::Xml(ref err) => (err as &error::Error).description(),
            Error::Xlsx(_) => "Xlsx operation fails",
        }
    }

    fn cause(&self) -> Option<&error::Error> {
        match *self {
            Error::Io(ref err) => (err as &error::Error).cause(),
            Error::Zip(ref err) => (err as &error::Error).cause(),
            Error::Xml(ref err) => (err as &error::Error).cause(),
            Error::Xlsx(_) => None,
        }
    }
}

pub type XlsxResult<T> = Result<T, Error>;
