#[derive(Debug, Deserialize)]
struct Theme {
    themeElements: ThemeElements,
    extraClrSchemeLst: Option<()>,
}

#[derive(Debug, Deserialize)]
struct ThemeElements {
    clrScheme: ClrScheme,
}

#[derive(Debug)]
struct ClrScheme {
    name: String,
    //#[serde(rename = "$value", default)]
    //clrs: Map<String, Clr>,
    dk1: Clr,
}

#[derive(Debug, Deserialize)]
enum Clr {
    #[serde(rename = "sysClr")]
    SysClr { val: String, lastClr: String },
    #[serde(rename = "srgbClr")]
    SrgbClr { val: String },
}


impl_from_xml_str!(Theme);

test_load_from_xml_str!(Theme, "tests/xlsx/xl/theme/theme1.xml");

#[allow(non_upper_case_globals, unused_attributes, unused_qualifications)]
const _IMPL_DESERIALIZE_FOR_ClrScheme: () = {
    extern crate serde as _serde;
    #[automatically_derived]
    impl<'de> _serde::Deserialize<'de> for ClrScheme {
        fn deserialize<__D>(__deserializer: __D) -> _serde::export::Result<Self, __D::Error>
        where
            __D: _serde::Deserializer<'de>,
        {
            #[allow(non_camel_case_types)]
            enum __Field {
                __field0,
                __field1,
                __ignore,
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
                fn visit_u64<__E>(self, __value: u64) -> _serde::export::Result<Self::Value, __E>
                where
                    __E: _serde::de::Error,
                {
                    match __value {
                        0u64 => _serde::export::Ok(__Field::__field0),
                        1u64 => _serde::export::Ok(__Field::__field1),
                        _ => _serde::export::Err(_serde::de::Error::invalid_value(
                            _serde::de::Unexpected::Unsigned(__value),
                            &"field index 0 <= i < 2",
                        )),
                    }
                }
                fn visit_str<__E>(self, __value: &str) -> _serde::export::Result<Self::Value, __E>
                where
                    __E: _serde::de::Error,
                {
                    match __value {
                        "name" => _serde::export::Ok(__Field::__field0),
                        "dk1" => _serde::export::Ok(__Field::__field1),
                        _ => _serde::export::Ok(__Field::__ignore),
                    }
                }
                fn visit_bytes<__E>(
                    self,
                    __value: &[u8],
                ) -> _serde::export::Result<Self::Value, __E>
                where
                    __E: _serde::de::Error,
                {
                    match __value {
                        b"name" => _serde::export::Ok(__Field::__field0),
                        b"dk1" => _serde::export::Ok(__Field::__field1),
                        _ => _serde::export::Ok(__Field::__ignore),
                    }
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
                fn visit_seq<__A>(
                    self,
                    mut __seq: __A,
                ) -> _serde::export::Result<Self::Value, __A::Error>
                where
                    __A: _serde::de::SeqAccess<'de>,
                {
                    let __field0 =
                        match match _serde::de::SeqAccess::next_element::<String>(&mut __seq) {
                            Result::Ok(val) => val,
                            Result::Err(err) => {
                                return Result::Err(From::from(err))
                            }
                        } {
                            Some(__value) => __value,
                            None => {
                                return _serde::export::Err(_serde::de::Error::invalid_length(
                                    0usize,
                                    &"tuple of 2 elements",
                                ));
                            }
                        };
                    let __field1 =
                        match match _serde::de::SeqAccess::next_element::<Clr>(&mut __seq) {
                            Result::Ok(val) => val,
                            Result::Err(err) => {
                                return Result::Err(From::from(err))
                            }
                        } {
                            Some(__value) => __value,
                            None => {
                                return _serde::export::Err(_serde::de::Error::invalid_length(
                                    1usize,
                                    &"tuple of 2 elements",
                                ));
                            }
                        };
                    _serde::export::Ok(ClrScheme {
                        name: __field0,
                        dk1: __field1,
                    })
                }
                #[inline]
                fn visit_map<__A>(
                    self,
                    mut __map: __A,
                ) -> _serde::export::Result<Self::Value, __A::Error>
                where
                    __A: _serde::de::MapAccess<'de>,
                {
                    let mut __field0: _serde::export::Option<String> = _serde::export::None;
                    let mut __field1: _serde::export::Option<Clr> = _serde::export::None;
                    while let _serde::export::Some(__key) =
                        match _serde::de::MapAccess::next_key::<__Field>(&mut __map) {
                            Result::Ok(val) => val,
                            Result::Err(err) => {
                                return Result::Err(From::from(err))
                            }
                        }
                    {
                        match __key {
                            __Field::__field0 => {
                                if _serde::export::Option::is_some(&__field0) {
                                    return _serde::export::Err(
                                        <__A::Error as _serde::de::Error>::duplicate_field(
                                            "name",
                                        ),
                                    );
                                }
                                __field0 = _serde::export::Some(
                                    match _serde::de::MapAccess::next_value::<String>(
                                        &mut __map,
                                    ) {
                                        Result::Ok(val) => val,
                                        Result::Err(err) => {
                                            return Result::Err(From::from(err))
                                        }
                                    },
                                );
                            }
                            __Field::__field1 => {
                                if _serde::export::Option::is_some(&__field1) {
                                    return _serde::export::Err(
                                        <__A::Error as _serde::de::Error>::duplicate_field(
                                            "dk1",
                                        ),
                                    );
                                }
                                __field1 = _serde::export::Some(
                                    match _serde::de::MapAccess::next_value::<Clr>(
                                        &mut __map,
                                    ) {
                                        Result::Ok(val) => val,
                                        Result::Err(err) => {
                                            return Result::Err(From::from(err))
                                        }
                                    },
                                );
                            }
                            _ => {
                                let _ =
                                            match _serde::de::MapAccess::next_value::<_serde::de::IgnoredAny>(&mut __map)
                                                {
                                                Result::Ok(val) =>
                                                val,
                                                Result::Err(err) =>
                                                {
                                                    return Result::Err(From::from(err))
                                                }
                                            };
                            }
                        }
                    }
                    let __field0 = match __field0 {
                        _serde::export::Some(__field0) => __field0,
                        _serde::export::None => {
                            match _serde::private::de::missing_field("name") {
                                Result::Ok(val) => val,
                                Result::Err(err) => {
                                    return Result::Err(From::from(err))
                                }
                            }
                        }
                    };
                    let __field1 = match __field1 {
                        _serde::export::Some(__field1) => __field1,
                        _serde::export::None => {
                            match _serde::private::de::missing_field("dk1") {
                                Result::Ok(val) => val,
                                Result::Err(err) => {
                                    return Result::Err(From::from(err))
                                }
                            }
                        }
                    };
                    _serde::export::Ok(ClrScheme {
                        name: __field0,
                        dk1: __field1,
                    })
                }
            }
            const FIELDS: &'static [&'static str] = &["name", "dk1"];
            _serde::Deserializer::deserialize_struct(
                __deserializer,
                "ClrScheme",
                FIELDS,
                __Visitor {
                    marker: _serde::export::PhantomData::<ClrScheme>,
                    lifetime: _serde::export::PhantomData,
                },
            )
        }
    }
};
