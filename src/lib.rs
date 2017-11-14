#[macro_use]
extern crate serde_derive;
extern crate serde;
extern crate serde_xml_rs;

mod xlsx_style;

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }
}
