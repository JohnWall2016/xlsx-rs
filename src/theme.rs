#[derive(Debug, Deserialize)]
struct Theme {
    themeElements: ThemeElements,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
struct ThemeElements {
    clrScheme: ClrScheme,
}

use std::collections::BTreeMap as Map;

#[derive(Debug)]
struct ClrScheme {
    name: String,
    clrs: Map<String, Option<Clr>>,
}

#[derive(Debug, Deserialize)]
enum Clr {
    #[serde(rename = "sysClr")]
    SysClr { val: String, lastClr: String },
    #[serde(rename = "srgbClr")]
    SrgbClr { val: String },
}

#[allow(non_upper_case_globals, unused_attributes, unused_qualifications)]
const _IMPL_DESERIALIZE_FOR_ClrScheme: () = {
    extern crate serde as _serde;
    use std::cell::RefCell;
    
    #[automatically_derived]
    impl<'de> _serde::Deserialize<'de> for ClrScheme {
        fn deserialize<__D>(__deserializer: __D) -> _serde::export::Result<Self, __D::Error>
        where
            __D: _serde::Deserializer<'de>,
        {
            #[allow(non_camel_case_types)]
            struct __Field {
                name: String,
            }
            struct __FieldVisitor;
            impl<'de> _serde::de::Visitor<'de> for __FieldVisitor {
                type Value = __Field;
                fn expecting(
                    &self,
                    formatter: &mut _serde::export::Formatter,
                ) -> _serde::export::fmt::Result {
                    _serde::export::Formatter::write_str(formatter, "field identifier")
                }
                fn visit_str<__E>(self, __value: &str) -> _serde::export::Result<Self::Value, __E>
                where
                    __E: _serde::de::Error,
                {
                    //println!("visit_str:{}", __value);
                    _serde::export::Ok(__Field { name: String::from(__value) })
                }
            }
            impl<'de> _serde::Deserialize<'de> for __Field {
                #[inline]
                fn deserialize<__D>(__deserializer: __D) -> _serde::export::Result<Self, __D::Error>
                where
                    __D: _serde::Deserializer<'de>,
                {
                    _serde::Deserializer::deserialize_identifier(__deserializer, __FieldVisitor)
                }
            }
            struct __Visitor<'de> {
                marker: _serde::export::PhantomData<ClrScheme>,
                lifetime: _serde::export::PhantomData<&'de ()>,
                clrs: RefCell<Map<String, Option<Clr>>>,
            }
            impl<'de> _serde::de::Visitor<'de> for __Visitor<'de> {
                type Value = ClrScheme;
                fn expecting(
                    &self,
                    formatter: &mut _serde::export::Formatter,
                ) -> _serde::export::fmt::Result {
                    _serde::export::Formatter::write_str(formatter, "struct ClrScheme")
                }
                #[inline]
                fn visit_map<__A>(
                    self,
                    mut __map: __A,
                ) -> _serde::export::Result<Self::Value, __A::Error>
                where
                    __A: _serde::de::MapAccess<'de>,
                {
                    //println!("visit_map:");
                    let mut __field0: _serde::export::Option<String> = _serde::export::None;
                    while let _serde::export::Some(__key) =
                        match _serde::de::MapAccess::next_key::<__Field>(&mut __map) {
                            Result::Ok(val) => val,
                            Result::Err(err) => return Result::Err(From::from(err)),
                        }
                    {
                        //println!("visit_map2:{}", __key.name);
                        if __key.name == "name" {
                            if _serde::export::Option::is_some(&__field0) {
                                return _serde::export::Err(
                                    <__A::Error as _serde::de::Error>::duplicate_field("name"),
                                );
                            }
                            __field0 =
                                _serde::export::Some(
                                    match _serde::de::MapAccess::next_value::<String>(&mut __map) {
                                        Result::Ok(val) => val,
                                        Result::Err(err) => return Result::Err(From::from(err)),
                                    },
                                );
                        } else {
                            let mut clr = _serde::export::Some(
                                match _serde::de::MapAccess::next_value::<Clr>(&mut __map) {
                                    Result::Ok(val) => val,
                                    Result::Err(err) => return Result::Err(From::from(err)),
                                },
                            );
                            //println!("clr: {:?}", clr);
                            self.clrs.borrow_mut().insert(__key.name, clr);
                        }

                    }
                    let __field0 = match __field0 {
                        _serde::export::Some(__field0) => __field0,
                        _serde::export::None => {
                            match _serde::private::de::missing_field("name") {
                                Result::Ok(val) => val,
                                Result::Err(err) => return Result::Err(From::from(err)),
                            }
                        }
                    };

                    _serde::export::Ok(ClrScheme {
                        name: __field0,
                        clrs: self.clrs.into_inner(),
                    })
                }
            }
            const FIELDS: &'static [&'static str] = &[];
            _serde::Deserializer::deserialize_struct(
                __deserializer,
                "ClrScheme",
                FIELDS,
                __Visitor {
                    marker: _serde::export::PhantomData::<ClrScheme>,
                    lifetime: _serde::export::PhantomData,
                    clrs: RefCell::new(Map::new()),
                },
            )
        }
    }
};


impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");
