//! # YaSerDe
//!
//! YaSerDe is a framework for ***ser***ializing and ***de***serializing Rust data
//! structures efficiently and generically from and into XML.

#[macro_use]
extern crate log;

#[cfg(feature = "yaserde_derive")]
#[allow(unused_imports)]
#[macro_use]
extern crate yaserde_derive;

use std::io::{Read, Write};
use xml::writer::XmlEvent;

pub mod de;
pub mod ser;

/// A **data structure** that can be deserialized from any data format supported by YaSerDe.
pub trait YaDeserialize: Sized {
  fn deserialize<R: Read>(reader: &mut de::Deserializer<R>) -> Result<Self, String>;
}

/// A **data structure** that can be serialized into any data format supported by YaSerDe.
pub trait YaSerialize: Sized {
  fn serialize<W: Write>(&self, writer: &mut ser::Serializer<W>) -> Result<(), String>;
}

/// A **visitor** that can be implemented to retrieve information from source file.
pub trait Visitor<'de>: Sized {
  /// The value produced by this visitor.
  type Value;

  fn visit_bool(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected bool {:?}", v))
  }

  fn visit_i8(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected i8 {:?}", v))
  }

  fn visit_u8(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected u8 {:?}", v))
  }

  fn visit_i16(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected i16 {:?}", v))
  }

  fn visit_u16(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected u16 {:?}", v))
  }

  fn visit_i32(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected i32 {:?}", v))
  }

  fn visit_u32(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected u32 {:?}", v))
  }

  fn visit_i64(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected i64 {:?}", v))
  }

  fn visit_u64(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected u64 {:?}", v))
  }

  fn visit_f32(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected f32 {:?}", v))
  }

  fn visit_f64(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected f64 {:?}", v))
  }

  fn visit_str(self, v: &str) -> Result<Self::Value, String> {
    Err(format!("Unexpected str {:?}", v))
  }
}

macro_rules! serialize_type {
  ($type:ty) => {
    impl YaSerialize for $type {
      fn serialize<W: Write>(&self, writer: &mut ser::Serializer<W>) -> Result<(), String> {
        let content = format!("{}", self);
        let event = XmlEvent::characters(&content);
        let _ret = writer.write(event);
        Ok(())
      }
    }
  };
}

serialize_type!(bool);
serialize_type!(char);

serialize_type!(usize);
serialize_type!(u8);
serialize_type!(u16);
serialize_type!(u32);
serialize_type!(u64);

serialize_type!(isize);
serialize_type!(i8);
serialize_type!(i16);
serialize_type!(i32);
serialize_type!(i64);

serialize_type!(f32);
serialize_type!(f64);

#[test]
fn default_visitor() {
  struct Test;
  impl<'de> Visitor<'de> for Test {
    type Value = u8;
  }

  macro_rules! test_type {
    ($visitor:tt, $message:expr) => {{
      let t = Test {};
      assert_eq!(t.$visitor(""), Err($message.to_string()));
    }};
  }

  test_type!(visit_bool, "Unexpected bool \"\"");
  test_type!(visit_i8, "Unexpected i8 \"\"");
  test_type!(visit_u8, "Unexpected u8 \"\"");
  test_type!(visit_i16, "Unexpected i16 \"\"");
  test_type!(visit_u16, "Unexpected u16 \"\"");
  test_type!(visit_i32, "Unexpected i32 \"\"");
  test_type!(visit_u32, "Unexpected u32 \"\"");
  test_type!(visit_i64, "Unexpected i64 \"\"");
  test_type!(visit_u64, "Unexpected u64 \"\"");
  test_type!(visit_str, "Unexpected str \"\"");
}

#[doc(hidden)]
mod testing {
  #[macro_export]
  macro_rules! test_for_type {
    ($type:ty, $value:expr, $content:expr) => {{
      #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
      #[yaserde(rename = "data")]
      pub struct Data {
        item: $type,
      }

      let model = Data { item: $value };

      let content = if let Some(str_value) = $content {
        String::from("<data><item>") + str_value + "</item></data>"
      } else {
        String::from("<data />")
      };

      serialize_and_validate!(model, content);
      deserialize_and_validate!(&content, model, Data);
    }};
  }

  #[macro_export]
  macro_rules! test_for_attribute_type {
    ($type: ty, $value: expr, $content: expr) => {{
      #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
      #[yaserde(rename = "data")]
      pub struct Data {
        #[yaserde(attribute)]
        item: $type,
      }
      let model = Data { item: $value };

      let content = if let Some(str_value) = $content {
        "<data item=\"".to_string() + str_value + "\" />"
      } else {
        "<data />".to_string()
      };

      serialize_and_validate!(model, content);
      deserialize_and_validate!(&content, model, Data);
    }};
  }

  #[macro_export]
  macro_rules! deserialize_and_validate {
    ($content: expr, $model: expr, $struct: tt) => {
      let loaded: Result<$struct, String> = yaserde::de::from_str($content);
      assert_eq!(loaded, Ok($model));
    };
  }

  #[macro_export]
  macro_rules! serialize_and_validate {
    ($model: expr, $content: expr) => {
      let data: Result<String, String> = yaserde::ser::to_string(&$model);

      let content = String::from(r#"<?xml version="1.0" encoding="utf-8"?>"#) + &$content;
      assert_eq!(
        data,
        Ok(
          String::from(content)
            .split("\n")
            .map(|s| s.trim())
            .collect::<String>()
        )
      );
    };
  }
}