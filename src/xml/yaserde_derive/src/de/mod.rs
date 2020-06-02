pub mod build_default_value;
pub mod expand_enum;
pub mod expand_struct;

use crate::common::YaSerdeAttribute;
use proc_macro2::TokenStream;
use syn::Ident;

pub fn expand_derive_deserialize(ast: &syn::DeriveInput) -> Result<TokenStream, String> {
  let name = &ast.ident;
  let attrs = &ast.attrs;
  let data = &ast.data;

  let root_attributes = YaSerdeAttribute::parse(attrs);

  let root_name = format!(
    "{}{}",
    root_attributes.prefix_namespace(),
    root_attributes.xml_element_name(name)
  );

  let impl_block = match *data {
    syn::Data::Struct(ref data_struct) => {
      expand_struct::parse(data_struct, name, &root_name, &root_attributes)
    }
    syn::Data::Enum(ref data_enum) => {
      expand_enum::parse(data_enum, name, &root_name, &root_attributes)
    }
    syn::Data::Union(ref _data_union) => unimplemented!(),
  };

  let dummy_const = Ident::new(&format!("_IMPL_YA_DESERIALIZE_FOR_{}", name), name.span());

  let generated = quote! {
    #[allow(non_upper_case_globals, unused_attributes, unused_qualifications)]
    const #dummy_const: () = {
      extern crate yaserde as _yaserde;
      #impl_block
    };
  };

  Ok(generated)
}
