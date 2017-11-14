#![allow(non_snake_case)]

#[macro_use]
extern crate serde_derive;
extern crate serde;
extern crate serde_xml_rs;

/// serde_xlsx_items_struct!
///
/// ```rust
/// serde_xlsx_items_struct!(
///     NumFmts,
///     "numFmt" => NumFmt,
/// );
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct NumFmts {
///     count: String,
///
///     #[serde(rename = "numFmt", default)]
///     items: Vec<NumFmt>,
/// }
/// ```
///
/// ```rust
/// serde_xlsx_items_struct!{
///     name: SharedString,
///     item: "si" => StringItem,
///     fields: {
///         uniqueCount: String,
///         ...
///     }
/// };
/// ```
/// generate:
/// ```rust
/// #[derive(Debug, Deserialize)]
/// struct SharedString {
///     count: String,
///
///     #[serde(rename = "si", default)]
///     items: Vec<StringItem>,
///     uniqueCount: String,
///     ...
/// }
/// ```
///
macro_rules! serde_xlsx_items_struct {
    ($struct_name:ident,
     $serde_name:tt  => $items_struct_name:ident
    ) => {
        serde_xlsx_items_struct!(
            name: $struct_name,
            item: $serde_name => $items_struct_name,
            fields: {}
        );
    };
    (name: $struct_name:ident,
     item: $serde_name:tt  => $items_struct_name:ident
    ) => {
        serde_xlsx_items_struct!(
            name: $struct_name,
            item: $serde_name => $items_struct_name,
            fields: {}
        );
    };
    (name: $struct_name:ident,
     item: $serde_name:tt  =>  $items_struct_name:ident,
     fields: {
         $($element: ident: $ty: ty),*
     }) => {
        #[derive(Debug, Deserialize)]
        struct $struct_name {
            count: String,

            #[serde(rename = $serde_name, default)]
            items: Vec<$items_struct_name>,

            $($element: $ty),*
        }
    };
}

macro_rules! impl_from_xml_str {
    ($struct_name:ident) => {
        use serde_xml_rs::{deserialize, Error};

        impl $struct_name {
            fn from_xml_str(str: &String) -> Result<$struct_name, Error> {
                deserialize(str.as_bytes())
            }
        }
    }
}

macro_rules! test_load_from_xml_str{
    ($struct_name:ident, $xml_file_path:tt) => {
        #[test]
        fn load_from_xlsx_str() {
            use std::io::prelude::*;
            use std::fs::File;

            let path = format!("{}/{}", env!("CARGO_MANIFEST_DIR"), $xml_file_path);
            match File::open(&path) {
                Ok(mut file) => {
                    let mut contents = String::new();
                    match file.read_to_string(&mut contents) {
                        Ok(_) => {
                            match $struct_name::from_xml_str(&contents) {
                                Ok(ss) => println!("{:#?}", ss),
                                Err(err) => println!("{:#?}", err),
                            }
                        }
                        Err(err) => println!("read file error: {}", err),
                    }
                }
                Err(err) => println!("open file error: {}", err),
            }
        }
    };
}

mod styles;
mod shared_strings;
mod workbook;
