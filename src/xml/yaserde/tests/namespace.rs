#[macro_use]
extern crate yaserde;
#[macro_use]
extern crate yaserde_derive;

use std::io::{Read, Write};
use yaserde::{YaDeserialize, YaSerialize};

#[test]
fn struct_simple_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "book",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain"
  )]
  pub struct Book {
    #[yaserde(prefix = "ns")]
    author: String,
    #[yaserde(prefix = "ns")]
    title: String,
  }

  let content = r#"
    <ns:book xmlns:ns="http://www.sample.com/ns/domain">
      <ns:author>Antoine de Saint-Exupéry</ns:author>
      <ns:title>Little prince</ns:title>
    </ns:book>
  "#;

  let model = Book {
    author: String::from("Antoine de Saint-Exupéry"),
    title: String::from("Little prince"),
  };

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, Book);
}

#[test]
fn struct_multiple_namespaces() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "book",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain",
    namespace = "ns2: http://www.sample.com/ns/domain_2"
  )]
  pub struct Book {
    #[yaserde(prefix = "ns")]
    author: String,
    #[yaserde(prefix = "ns2")]
    title: String,
  }

  let content = r#"
    <ns:book xmlns:ns="http://www.sample.com/ns/domain" xmlns:ns2="http://www.sample.com/ns/domain_2">
      <ns:author>Antoine de Saint-Exupéry</ns:author>
      <ns2:title>Little prince</ns2:title>
    </ns:book>
  "#;

  let model = Book {
    author: String::from("Antoine de Saint-Exupéry"),
    title: String::from("Little prince"),
  };

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, Book);
}

#[test]
fn struct_partial_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "book",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain"
  )]
  pub struct Book {
    author: String,
    #[yaserde(prefix = "ns")]
    title: String,
  }

  let content = r#"
    <ns:book xmlns:ns="http://www.sample.com/ns/domain">
      <author>Antoine de Saint-Exupéry</author>
      <ns:title>Little prince</ns:title>
    </ns:book>
  "#;

  let model = Book {
    author: String::from("Antoine de Saint-Exupéry"),
    title: String::from("Little prince"),
  };

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, Book);
}

#[test]
fn struct_sub_namespace_definition() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "book",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain",
    namespace = "ns2: http://www.sample.com/ns/domain_2"
  )]
  pub struct Book {
    #[yaserde(prefix = "ns")]
    author: String,
    #[yaserde(prefix = "ns2", namespace = "ns2: http://www.sample.com/ns/domain_2")]
    title: String,
  }

  let content = r#"
    <ns:book xmlns:ns="http://www.sample.com/ns/domain">
      <ns:author>Antoine de Saint-Exupéry</ns:author>
      <ns2:title xmlns:ns2="http://www.sample.com/ns/domain_2">Little prince</ns2:title>
    </ns:book>
  "#;

  let model = Book {
    author: String::from("Antoine de Saint-Exupéry"),
    title: String::from("Little prince"),
  };

  // TODO support namespace for attribute to specify local namespace
  // serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, Book);
}

#[test]
fn struct_namespace_nested() {
  #[derive(Debug, Default, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(prefix = "nsa", namespace = "nsa: http://www.sample.com/ns/a")]
  struct A {
    #[yaserde(prefix = "nsa")]
    alpha: i32,
  }

  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(prefix = "nsb", namespace = "nsb: http://www.sample.com/ns/b")]
  struct B {
    // Note that name `nested` resides in `nsb` though it has a type from `nsa`
    #[yaserde(prefix = "nsb")]
    nested: A,
  }

  let content = r#"
    <nsb:B xmlns:nsb="http://www.sample.com/ns/b">
      <nsb:nested xmlns:nsa="http://www.sample.com/ns/a">
        <nsa:alpha>32</nsa:alpha>
      </nsb:nested>
    </nsb:B>
  "#;

  let model = B {
    nested: A { alpha: 32 },
  };

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, B);
}

#[test]
fn struct_namespace_nested_defined_at_root() {
  #[derive(Debug, Default, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(prefix = "nsa", namespace = "nsa: http://www.sample.com/ns/a")]
  struct A {
    #[yaserde(prefix = "nsa")]
    alpha: i32,
  }

  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    prefix = "nsb",
    namespace = "nsb: http://www.sample.com/ns/b"
    namespace = "nsa: http://www.sample.com/ns/a"
  )]
  struct B {
    // Note that name `nested` resides in `nsb` though it has a type from `nsa`
    #[yaserde(prefix = "nsb")]
    nested: A,
  }

  let content = r#"
    <nsb:B xmlns:nsa="http://www.sample.com/ns/a" xmlns:nsb="http://www.sample.com/ns/b">
      <nsb:nested>
        <nsa:alpha>32</nsa:alpha>
      </nsb:nested>
    </nsb:B>
  "#;

  let model = B {
    nested: A { alpha: 32 },
  };

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, B);
}

#[test]
fn struct_attribute_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "root",
    namespace = "ns1: http://www.sample.com/ns/domain1",
    namespace = "ns2: http://www.sample.com/ns/domain2"
  )]
  pub struct XmlStruct {
    #[yaserde(prefix = "ns1")]
    item_1: String,
    #[yaserde(attribute, prefix = "ns2")]
    item_2: String,
  }

  let model = XmlStruct {
    item_1: "something 1".to_string(),
    item_2: "something 2".to_string(),
  };

  let content = r#"
    <root xmlns:ns1="http://www.sample.com/ns/domain1" xmlns:ns2="http://www.sample.com/ns/domain2" ns2:item_2="something 2">
      <ns1:item_1>something 1</ns1:item_1>
    </root>
  "#;

  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn struct_implicit_default_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "tt",
    namespace = "http://www.w3.org/ns/ttml",
    namespace = "ttm: http://www.w3.org/ns/ttml#metadata"
  )]
  pub struct XmlStruct {
    item: String,
  }

  let model = XmlStruct {
    item: "something".to_string(),
  };

  let content = r#"<tt xmlns="http://www.w3.org/ns/ttml" xmlns:ttm="http://www.w3.org/ns/ttml#metadata"><item>something</item></tt>"#;
  serialize_and_validate!(model, content);
  // TODO
  // deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn struct_explicit_default_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "tt",
    default_namespace = "ttml",
    namespace = "ttml: http://www.w3.org/ns/ttml",
    namespace = "ttm: http://www.w3.org/ns/ttml#metadata"
  )]
  pub struct XmlStruct {
    item: String,
  }

  let model = XmlStruct {
    item: "something".to_string(),
  };

  let content = r#"<tt xmlns="http://www.w3.org/ns/ttml" xmlns:ttm="http://www.w3.org/ns/ttml#metadata"><item>something</item></tt>"#;
  serialize_and_validate!(model, content);
  // TODO
  // deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn struct_default_namespace_via_attribute_with_prefix() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "tt",
    prefix = "TTML",
    default_namespace = "TTML",
    namespace = "TTML: http://www.w3.org/ns/ttml",
    namespace = "ttm: http://www.w3.org/ns/ttml#metadata"
  )]
  pub struct XmlStruct {
    #[yaserde(prefix = "TTML")]
    item: String,
  }

  let model = XmlStruct {
    item: "something".to_string(),
  };

  let content = r#"<tt xmlns="http://www.w3.org/ns/ttml" xmlns:ttm="http://www.w3.org/ns/ttml#metadata"><item>something</item></tt>"#;
  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn enum_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "root",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain"
  )]
  pub enum XmlStruct {
    #[yaserde(prefix = "ns")]
    Item,
  }

  impl Default for XmlStruct {
    fn default() -> XmlStruct {
      XmlStruct::Item
    }
  }

  let content = r#"
    <ns:root xmlns:ns="http://www.sample.com/ns/domain">
      ns:Item
    </ns:root>
  "#;

  let model = XmlStruct::Item;
  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn enum_multi_namespaces() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "root",
    namespace = "ns1: http://www.sample.com/ns/domain1",
    namespace = "ns2: http://www.sample.com/ns/domain2"
  )]
  pub enum XmlStruct {
    #[yaserde(prefix = "ns1")]
    Item1,
    #[yaserde(prefix = "ns2")]
    Item2,
  }

  impl Default for XmlStruct {
    fn default() -> XmlStruct {
      XmlStruct::Item1
    }
  }

  let model = XmlStruct::Item1;
  let content = r#"
    <root xmlns:ns1="http://www.sample.com/ns/domain1" xmlns:ns2="http://www.sample.com/ns/domain2">
      ns1:Item1
    </root>
  "#;
  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);

  let model = XmlStruct::Item2;
  let content = r#"
    <root xmlns:ns1="http://www.sample.com/ns/domain1" xmlns:ns2="http://www.sample.com/ns/domain2">
      ns2:Item2
    </root>
  "#;
  serialize_and_validate!(model, content);
  // TODO
  // deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn enum_attribute_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "rootA",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain"
  )]
  pub enum XmlStruct {
    #[yaserde(prefix = "ns")]
    Item,
    #[yaserde(prefix = "ns")]
    ItemWithField(String),
  }

  impl Default for XmlStruct {
    fn default() -> XmlStruct {
      XmlStruct::Item
    }
  }

  let content = r#"
    <ns:rootA xmlns:ns="http://www.sample.com/ns/domain">
      ns:Item
    </ns:rootA>
  "#;

  let model = XmlStruct::Item;
  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);

  let model = XmlStruct::ItemWithField("Value".to_string());

  let content = r#"<ns:rootA xmlns:ns="http://www.sample.com/ns/domain"><ns:ItemWithField>Value</ns:ItemWithField></ns:rootA>"#;
  serialize_and_validate!(model, content);
  deserialize_and_validate!(content, model, XmlStruct);
}

#[test]
fn struct_bad_namespace() {
  #[derive(Debug, PartialEq, YaDeserialize, YaSerialize)]
  #[yaserde(
    rename = "book",
    prefix = "ns",
    namespace = "ns: http://www.sample.com/ns/domain",
    namespace = "ns2: http://www.sample.com/ns/domain_2"
  )]
  pub struct Book {
    #[yaserde(prefix = "ns")]
    author: String,
    #[yaserde(prefix = "ns2", namespace = "ns2: http://www.sample.com/ns/domain_2")]
    title: String,
  }

  let content = r#"
    <ns:book xmlns:ns="http://www.sample.com/ns/domain2">
      <ns:author>Antoine de Saint-Exupéry</ns:author>
      <ns:title>Little prince</ns:title>
    </ns:book>
  "#;

  let loaded: Result<Book, String> = yaserde::de::from_str(content);
  assert_eq!(
    loaded,
    Err("bad namespace for book, found http://www.sample.com/ns/domain2".to_string())
  );
}
