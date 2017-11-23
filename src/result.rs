
use std::io;
//extern crate zip;
use zip;

#[derive(Debug)]
pub enum Error {
    Io(io::Error),
    Zip(zip::result::ZipError),
    Xlsx(String),
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
            Error::Xlsx(_) => "Xlsx operation fails",
        }
    }

    fn cause(&self) -> Option<&error::Error> {
        match *self {
            Error::Io(ref err) => (err as &error::Error).cause(),
            Error::Zip(ref err) => (err as &error::Error).cause(),
            Error::Xlsx(_) => None,
        }
    }
}

pub type XlsxResult<T> = Result<T, Error>;
