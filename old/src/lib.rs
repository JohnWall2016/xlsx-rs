#![allow(non_snake_case, dead_code, unused_macros)]

#[macro_use] extern crate serde_derive;
//extern crate serde;
extern crate serde_xml_rs;
#[macro_use] extern crate maplit;

mod xml;

#[cfg(feature = "bzip2")]
extern crate bzip2;
extern crate flate2;
extern crate msdos_time;
extern crate podio;
extern crate time;

mod zip;

mod file;
mod refer;
mod xlsx;

mod result;

