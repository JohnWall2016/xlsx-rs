use crate::common::attribute::YaSerdeAttribute;
use heck::CamelCase;
use proc_macro2::Span;
use proc_macro2::{Ident, TokenStream};
use std::fmt;
use syn::spanned::Spanned;
use syn::Type::Path;

#[derive(Debug)]
pub struct YaSerdeField {
  syn_field: syn::Field,
  attributes: YaSerdeAttribute,
}

impl YaSerdeField {
  pub fn new(syn_field: syn::Field) -> Self {
    let attributes = YaSerdeAttribute::parse(&syn_field.attrs);

    YaSerdeField {
      syn_field,
      attributes,
    }
  }

  pub fn is_attribute(&self) -> bool {
    self.attributes.attribute
  }

  pub fn is_text_content(&self) -> bool {
    self.attributes.text
  }

  pub fn is_flatten(&self) -> bool {
    self.attributes.flatten
  }

  pub fn label(&self) -> Option<Ident> {
    self.syn_field.ident.clone()
  }

  pub fn get_value_label(&self) -> Option<syn::Ident> {
    self
      .syn_field
      .ident
      .clone()
      .map(|ident| syn::Ident::new(&format!("__{}_value", ident.to_string()), ident.span()))
  }

  pub fn renamed_label_without_namespace(&self) -> String {
    self
      .attributes
      .rename
      .clone()
      .unwrap_or_else(|| self.label().as_ref().unwrap().to_string())
  }

  pub fn get_match_arm_from_rename(&self) -> Vec<String> {
    if let Some(rename) = &self.attributes.rename {
      rename
        .split("|")
        .filter_map(|s| {
          let r = s.trim();
          if r.is_empty() {
            None
          } else {
            Some(r.to_string())
          }
        })
        .collect()
    } else {
      vec![self.label().as_ref().unwrap().to_string()]
    }
  }

  pub fn renamed_label(&self, root_attributes: &YaSerdeAttribute) -> String {
    let prefix = if root_attributes.default_namespace == self.attributes.prefix {
      "".to_string()
    } else {
      self
        .attributes
        .prefix
        .clone()
        .map_or("".to_string(), |prefix| prefix + ":")
    };

    let label = self.renamed_label_without_namespace();

    format!("{}{}", prefix, label)
  }

  pub fn get_visitor_ident(&self, struct_name: Option<&syn::Path>) -> Ident {
    let label = self.renamed_label_without_namespace();

    let struct_id = struct_name.map_or_else(
      || "".to_string(),
      |struct_name| {
        struct_name
          .segments
          .iter()
          .map(|s| s.ident.to_string())
          .collect()
      },
    );

    Ident::new(
      &format!(
        "__Visitor_{}_{}",
        label.replace(".", "_").to_camel_case(),
        struct_id
      ),
      self.get_span(),
    )
  }

  pub fn get_type(&self) -> Field {
    Field::from(&self.syn_field)
  }

  pub fn get_span(&self) -> Span {
    self.syn_field.span()
  }

  pub fn get_default_function(&self) -> Option<Ident> {
    self
      .attributes
      .default
      .as_ref()
      .map(|default| Ident::new(&default, self.get_span()))
  }

  pub fn get_skip_serializing_if_function(&self) -> Option<Ident> {
    self
      .attributes
      .skip_serializing_if
      .as_ref()
      .map(|skip_serializing_if| Ident::new(&skip_serializing_if, self.get_span()))
  }

  pub fn get_namespace_matching(
    &self,
    root_attributes: &YaSerdeAttribute,
    element_namespace: TokenStream,
    element_name: TokenStream,
  ) -> TokenStream {
    root_attributes.get_namespace_matching(
      &self.attributes.prefix,
      element_namespace,
      element_name,
      false,
    )
  }

  pub fn ser_wrap_default_attribute(
    &self,
    builder: Option<TokenStream>,
    setter: TokenStream,
  ) -> TokenStream {
    let label = self.label();

    let yaserde_inner_definition = builder
      .map(|builder| quote!(let yaserde_inner = #builder;))
      .unwrap_or(quote!());

    self
      .get_default_function()
      .map(|default_function| {
        quote! {
          #yaserde_inner_definition
          let struct_start_event =
            if self.#label != #default_function() {
              #setter
            } else {
              struct_start_event
            };
        }
      })
      .unwrap_or(quote! {
        #yaserde_inner_definition
        let struct_start_event = #setter;
      })
  }
}

#[derive(Debug)]
pub enum Field {
  FieldString,
  FieldBool,
  FieldI8,
  FieldU8,
  FieldI16,
  FieldU16,
  FieldI32,
  FieldU32,
  FieldI64,
  FieldU64,
  FieldF32,
  FieldF64,
  FieldOption { data_type: Box<Field> },
  FieldVec { data_type: Box<Field> },
  FieldStruct { struct_name: syn::Path },
}

impl Field {
  pub fn get_simple_type_visitor(&self) -> TokenStream {
    let ident = format_ident!("visit_{}", self.to_string());
    quote! {#ident}
  }
}

impl From<&syn::Path> for Field {
  fn from(path: &syn::Path) -> Self {
    let result = if let Some(segment) = path.segments.last() {
      match segment.ident.to_string().as_str() {
        "String" => Some(Field::FieldString),
        "bool" => Some(Field::FieldBool),
        "i8" => Some(Field::FieldI8),
        "u8" => Some(Field::FieldU8),
        "i16" => Some(Field::FieldI16),
        "u16" => Some(Field::FieldU16),
        "i32" => Some(Field::FieldI32),
        "u32" => Some(Field::FieldU32),
        "i64" => Some(Field::FieldI64),
        "u64" => Some(Field::FieldU64),
        "f32" => Some(Field::FieldF32),
        "f64" => Some(Field::FieldF64),
        "Option" => Some(Field::FieldOption {
          data_type: Box::new(Field::from(segment)),
        }),
        "Vec" => Some(Field::FieldVec {
          data_type: Box::new(Field::from(segment)),
        }),
        _ => None,
      }
    } else {
      None
    };

    result.unwrap_or_else(|| Field::FieldStruct {
      struct_name: path.clone(),
    })
  }
}

impl From<&syn::Field> for Field {
  fn from(field: &syn::Field) -> Self {
    let mut ty = &field.ty;
    while let syn::Type::Group(g) = ty {
      ty = &g.elem;
    }
    match ty {
      Path(ref path) => Field::from(&path.path),
      _ => panic!("unable to match {:?}", field.ty),
    }
  }
}

impl From<&syn::PathSegment> for Field {
  fn from(path_segment: &syn::PathSegment) -> Self {
    if let syn::PathArguments::AngleBracketed(ref args) = path_segment.arguments {
      if let Some(tt) = args.args.first() {
        if let syn::GenericArgument::Type(ref argument) = *tt {
          if let Path(ref path) = *argument {
            return Field::from(&path.path);
          }
        }
      }
    }
    unimplemented!()
  }
}

impl Into<proc_macro2::TokenStream> for Field {
  fn into(self) -> proc_macro2::TokenStream {
    match self {
      Field::FieldString => quote! {String},
      Field::FieldBool => quote! {bool},
      Field::FieldI8 => quote! {i8},
      Field::FieldU8 => quote! {u8},
      Field::FieldI16 => quote! {i16},
      Field::FieldU16 => quote! {u16},
      Field::FieldI32 => quote! {i32},
      Field::FieldU32 => quote! {u32},
      Field::FieldI64 => quote! {i64},
      Field::FieldU64 => quote! {u64},
      Field::FieldF32 => quote! {f32},
      Field::FieldF64 => quote! {f64},
      _ => panic!("Not a simple type: {:?}", self),
    }
  }
}

impl Into<String> for &Field {
  fn into(self) -> String {
    match self {
      Field::FieldString => "str".to_string(),
      Field::FieldBool => "bool".to_string(),
      Field::FieldI8 => "i8".to_string(),
      Field::FieldU8 => "u8".to_string(),
      Field::FieldI16 => "i16".to_string(),
      Field::FieldU16 => "u16".to_string(),
      Field::FieldI32 => "i32".to_string(),
      Field::FieldU32 => "u32".to_string(),
      Field::FieldI64 => "i64".to_string(),
      Field::FieldU64 => "u64".to_string(),
      Field::FieldF32 => "f32".to_string(),
      Field::FieldF64 => "f64".to_string(),
      _ => panic!("Not a simple type: {:?}", self),
    }
  }
}

impl fmt::Display for Field {
  fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
    let string_representation: String = self.into();
    write!(f, "{}", string_representation)
  }
}
