use crate::common::{Field, YaSerdeAttribute, YaSerdeField};

use crate::ser::{element::*, implement_deserializer::implement_deserializer};
use proc_macro2::TokenStream;
use syn::DataStruct;
use syn::Ident;

pub fn serialize(
  data_struct: &DataStruct,
  name: &Ident,
  root: &str,
  root_attributes: &YaSerdeAttribute,
) -> TokenStream {
  let build_attributes: TokenStream = data_struct
    .fields
    .iter()
    .map(|field| YaSerdeField::new(field.clone()))
    .filter(|field| field.is_attribute())
    .map(|field| {
      let label = field.label();
      let label_name = field.renamed_label(root_attributes);

      match field.get_type() {
        Field::FieldString
        | Field::FieldBool
        | Field::FieldI8
        | Field::FieldU8
        | Field::FieldI16
        | Field::FieldU16
        | Field::FieldI32
        | Field::FieldU32
        | Field::FieldI64
        | Field::FieldU64
        | Field::FieldF32
        | Field::FieldF64 => Some(field.ser_wrap_default_attribute(
          Some(quote!(self.#label.to_string())),
          quote!({
            struct_start_event.attr(#label_name, &yaserde_inner)
          }),
        )),
        Field::FieldOption { data_type } => match *data_type {
          Field::FieldString => Some(field.ser_wrap_default_attribute(
            None,
            quote!({
              if let Some(ref value) = self.#label {
                struct_start_event.attr(#label_name, value)
              } else {
                struct_start_event
              }
            }),
          )),
          Field::FieldBool
          | Field::FieldI8
          | Field::FieldU8
          | Field::FieldI16
          | Field::FieldU16
          | Field::FieldI32
          | Field::FieldU32
          | Field::FieldI64
          | Field::FieldU64
          | Field::FieldF32
          | Field::FieldF64 => Some(field.ser_wrap_default_attribute(
            Some(quote!(self.#label.map_or_else(|| String::new(), |v| v.to_string()))),
            quote!({
              if let Some(ref value) = self.#label {
                struct_start_event.attr(#label_name, &yaserde_inner)
              } else {
                struct_start_event
              }
            }),
          )),
          Field::FieldVec { .. } => {
            let item_ident = Ident::new("yaserde_item", field.get_span());
            let inner = enclose_formatted_characters(&item_ident, label_name);

            Some(field.ser_wrap_default_attribute(
              None,
              quote!({
                if let Some(ref yaserde_list) = self.#label {
                  for yaserde_item in yaserde_list.iter() {
                    #inner
                  }
                }
              }),
            ))
          }
          Field::FieldStruct { .. } => Some(field.ser_wrap_default_attribute(
            Some(quote!(self.#label
                  .as_ref()
                  .map_or_else(|| Ok(String::new()), |v| yaserde::ser::to_string_content(v))?)),
            quote!({
              if let Some(ref yaserde_struct) = self.#label {
                struct_start_event.attr(#label_name, &yaserde_inner)
              } else {
                struct_start_event
              }
            }),
          )),
          Field::FieldOption { .. } => unimplemented!(),
        },
        Field::FieldStruct { .. } => Some(field.ser_wrap_default_attribute(
          Some(quote!(yaserde::ser::to_string_content(&self.#label)?)),
          quote!({
            struct_start_event.attr(#label_name, &yaserde_inner)
          }),
        )),
        Field::FieldVec { .. } => None,
      }
    })
    .filter_map(|x| x)
    .collect();

  let struct_inspector: TokenStream = data_struct
    .fields
    .iter()
    .map(|field| YaSerdeField::new(field.clone()))
    .filter(|field| !field.is_attribute())
    .map(|field| {
      let label = field.label();
      if field.is_text_content() {
        return Some(quote!(
          let data_event = XmlEvent::characters(&self.#label);
          writer.write(data_event).map_err(|e| e.to_string())?;
        ));
      }

      let label_name = field.renamed_label(root_attributes);
      let conditions = condition_generator(&label, &field);

      match field.get_type() {
        Field::FieldString
        | Field::FieldBool
        | Field::FieldI8
        | Field::FieldU8
        | Field::FieldI16
        | Field::FieldU16
        | Field::FieldI32
        | Field::FieldU32
        | Field::FieldI64
        | Field::FieldU64
        | Field::FieldF32
        | Field::FieldF64 => serialize_element(&label, label_name, &conditions),

        Field::FieldOption { data_type } => match *data_type {
          Field::FieldString
          | Field::FieldBool
          | Field::FieldI8
          | Field::FieldU8
          | Field::FieldI16
          | Field::FieldU16
          | Field::FieldI32
          | Field::FieldU32
          | Field::FieldI64
          | Field::FieldU64
          | Field::FieldF32
          | Field::FieldF64 => {
            let item_ident = Ident::new("yaserde_item", field.get_span());
            let inner = enclose_formatted_characters_for_value(&item_ident, label_name);

            Some(quote! {
              #conditions {
                if let Some(ref yaserde_item) = self.#label {
                  #inner
                }
              }
            })
          }
          Field::FieldVec { .. } => {
            let item_ident = Ident::new("yaserde_item", field.get_span());
            let inner = enclose_formatted_characters_for_value(&item_ident, label_name);

            Some(quote! {
              #conditions {
                if let Some(ref yaserde_items) = &self.#label {
                  for yaserde_item in yaserde_items.iter() {
                    #inner
                  }
                }
              }
            })
          }
          Field::FieldStruct { .. } => Some(if field.is_flatten() {
            quote! {
              if let Some(ref item) = &self.#label {
                writer.set_start_event_name(None);
                writer.set_skip_start_end(true);
                item.serialize(writer)?;
              }
            }
          } else {
            quote! {
              if let Some(ref item) = &self.#label {
                writer.set_start_event_name(Some(#label_name.to_string()));
                writer.set_skip_start_end(false);
                item.serialize(writer)?;
              }
            }
          }),
          _ => unimplemented!(),
        },
        Field::FieldStruct { .. } => {
          let (start_event, skip_start) = if field.is_flatten() {
            (quote!(None), true)
          } else {
            (quote!(Some(#label_name.to_string())), false)
          };

          Some(quote! {
            writer.set_start_event_name(#start_event);
            writer.set_skip_start_end(#skip_start);
            self.#label.serialize(writer)?;
          })
        }
        Field::FieldVec { data_type } => match *data_type {
          Field::FieldString => {
            let item_ident = Ident::new("yaserde_item", field.get_span());
            let inner = enclose_formatted_characters_for_value(&item_ident, label_name);

            Some(quote! {
              for yaserde_item in &self.#label {
                #inner
              }
            })
          }
          Field::FieldBool
          | Field::FieldI8
          | Field::FieldU8
          | Field::FieldI16
          | Field::FieldU16
          | Field::FieldI32
          | Field::FieldU32
          | Field::FieldI64
          | Field::FieldU64
          | Field::FieldF32
          | Field::FieldF64 => {
            let item_ident = Ident::new("yaserde_item", field.get_span());
            let inner = enclose_formatted_characters_for_value(&item_ident, label_name);

            Some(quote! {
              for yaserde_item in &self.#label {
                #inner
              }
            })
          }
          Field::FieldOption { .. } => Some(quote! {
            for item in &self.#label {
              if let Some(value) = item {
                writer.set_start_event_name(None);
                writer.set_skip_start_end(false);
                value.serialize(writer)?;
              }
            }
          }),
          Field::FieldStruct { .. } => Some(quote! {
            for item in &self.#label {
              writer.set_start_event_name(None);
              writer.set_skip_start_end(false);
              item.serialize(writer)?;
            }
          }),
          Field::FieldVec { .. } => {
            unimplemented!();
          }
        },
      }
    })
    .filter_map(|x| x)
    .collect();

  implement_deserializer(
    name,
    root,
    root_attributes,
    build_attributes,
    struct_inspector,
  )
}
