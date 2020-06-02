#[macro_use]
extern crate yaserde;
#[macro_use]
extern crate yaserde_derive;

use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

#[test]
fn basic_option_types() {
  test_for_type!(Option::<String>, Some("test".to_string()), Some("test"));
  test_for_type!(Option::<String>, None, None);
  test_for_type!(Option::<bool>, Some(true), Some("true"));
  test_for_type!(Option::<bool>, None, None);
  test_for_type!(Option::<u8>, Some(12 as u8), Some("12"));
  test_for_type!(Option::<u8>, None, None);
  test_for_type!(Option::<i8>, Some(12 as i8), Some("12"));
  test_for_type!(Option::<i8>, Some(-12 as i8), Some("-12"));
  test_for_type!(Option::<i8>, None, None);
  test_for_type!(Option::<u16>, Some(12 as u16), Some("12"));
  test_for_type!(Option::<u16>, None, None);
  test_for_type!(Option::<i16>, Some(12 as i16), Some("12"));
  test_for_type!(Option::<i16>, Some(-12 as i16), Some("-12"));
  test_for_type!(Option::<i16>, None, None);
  test_for_type!(Option::<u32>, Some(12 as u32), Some("12"));
  test_for_type!(Option::<u32>, None, None);
  test_for_type!(Option::<i32>, Some(12 as i32), Some("12"));
  test_for_type!(Option::<i32>, Some(-12 as i32), Some("-12"));
  test_for_type!(Option::<i32>, None, None);
  test_for_type!(Option::<u64>, Some(12 as u64), Some("12"));
  test_for_type!(Option::<u64>, None, None);
  test_for_type!(Option::<i64>, Some(12 as i64), Some("12"));
  test_for_type!(Option::<i64>, Some(-12 as i64), Some("-12"));
  test_for_type!(Option::<i64>, None, None);
  test_for_type!(Option::<f32>, Some(-12.5 as f32), Some("-12.5"));
  test_for_type!(Option::<f32>, None, None);
  test_for_type!(Option::<f64>, Some(-12.5 as f64), Some("-12.5"));
  test_for_type!(Option::<f64>, None, None);

  // TODO
  // test_for_type!(Option::<Vec::<u8>>, None, None);
  // test_for_type!(Option::<Vec::<u8>>, Some(vec![0]), Some("0"));
  // test_for_type!(Option::<Vec::<String>>, None, None);
  // test_for_type!(Option::<Vec::<String>>, Some(vec!["test".to_string()]), Some("test"));

  test_for_attribute_type!(Option::<String>, Some("test".to_string()), Some("test"));
  test_for_attribute_type!(Option::<String>, None, None);
  test_for_attribute_type!(Option::<bool>, Some(true), Some("true"));
  test_for_attribute_type!(Option::<bool>, None, None);
  test_for_attribute_type!(Option::<u8>, Some(12 as u8), Some("12"));
  test_for_attribute_type!(Option::<u8>, None, None);
  test_for_attribute_type!(Option::<i8>, Some(12 as i8), Some("12"));
  test_for_attribute_type!(Option::<i8>, Some(-12 as i8), Some("-12"));
  test_for_attribute_type!(Option::<i8>, None, None);
  test_for_attribute_type!(Option::<u16>, Some(12 as u16), Some("12"));
  test_for_attribute_type!(Option::<u16>, None, None);
  test_for_attribute_type!(Option::<i16>, Some(12 as i16), Some("12"));
  test_for_attribute_type!(Option::<i16>, Some(-12 as i16), Some("-12"));
  test_for_attribute_type!(Option::<i16>, None, None);
  test_for_attribute_type!(Option::<u32>, Some(12 as u32), Some("12"));
  test_for_attribute_type!(Option::<u32>, None, None);
  test_for_attribute_type!(Option::<i32>, Some(12 as i32), Some("12"));
  test_for_attribute_type!(Option::<i32>, Some(-12 as i32), Some("-12"));
  test_for_attribute_type!(Option::<i32>, None, None);
  test_for_attribute_type!(Option::<u64>, Some(12 as u64), Some("12"));
  test_for_attribute_type!(Option::<u64>, None, None);
  test_for_attribute_type!(Option::<i64>, Some(12 as i64), Some("12"));
  test_for_attribute_type!(Option::<i64>, Some(-12 as i64), Some("-12"));
  test_for_attribute_type!(Option::<i64>, None, None);
  test_for_attribute_type!(Option::<f32>, Some(-12.5 as f32), Some("-12.5"));
  test_for_attribute_type!(Option::<f32>, None, None);
  test_for_attribute_type!(Option::<f64>, Some(-12.5 as f64), Some("-12.5"));
  test_for_attribute_type!(Option::<f64>, None, None);
}

#[test]
fn option_struct() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  struct Test {
    field: SubTest,
  }

  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  struct SubTest {
    content: Option<String>,
  }

  impl Default for SubTest {
    fn default() -> Self {
      SubTest { content: None }
    }
  }

  test_for_type!(
    Option::<Test>,
    Some(Test {
      field: SubTest {
        content: Some("value".to_string())
      }
    }),
    Some("<field><content>value</content></field>")
  );
  test_for_type!(Option::<Test>, None, None);
}
