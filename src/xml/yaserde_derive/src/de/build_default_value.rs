use crate::common::YaSerdeField;
use proc_macro2::TokenStream;

pub fn build_default_value(
  field: &YaSerdeField,
  field_type: Option<TokenStream>,
  value: TokenStream,
) -> Option<TokenStream> {
  let label = field.get_value_label();

  let default_value = field
    .get_default_function()
    .map(|default_function| quote!(#default_function()))
    .unwrap_or_else(|| quote!(#value));

  let field_type = field_type
    .map(|field_type| quote!(: #field_type))
    .unwrap_or(quote!());

  Some(quote! {
    #[allow(unused_mut)]
    let mut #label #field_type = #default_value;
  })
}
