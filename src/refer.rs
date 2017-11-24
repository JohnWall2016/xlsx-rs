use std::collections::BTreeMap as Map;

#[derive(Debug)]
pub struct Strings {
    values: Vec<String>,
    index_map: Map<String, usize>,
}

impl Strings {
    pub fn new() -> Self {
        Strings {
            values: Vec::new(),
            index_map: Map::new(),
        }
    }

    pub fn add(&mut self, str: &String) -> usize {
        match self.index_map.get(str) {
            Some(&i) => i,
            None => {
                self.values.push(str.clone());
                let i = self.values.len() - 1;
                self.index_map.insert(str.clone(), i);
                i
            }
        }
    }

    pub fn index(&self, index: usize) -> Option<&String> {
        if index >= self.values.len() {
            None
        } else {
            Some(&self.values[index])
        }
    }

    pub fn get_index(&self, str: &String) -> Option<&usize> {
        self.index_map.get(str)
    }
}

#[derive(Debug)]
pub struct Colors {
    values: Map<String, String>,
}

impl Colors {
    pub fn new() -> Self {
        Colors { values: Map::new() }
    }

    pub fn insert(&mut self, name: &str, rgb_color: &String) -> Option<String> {
        self.values.insert(String::from(name), rgb_color.clone())
    }

    pub fn get(&self, name: &String) -> Option<&String> {
        self.values.get(name)
    }
}

#[derive(Debug)]
pub struct NumFmts {
    builtin_numfmts: Map<String, String>,
    defined_numfmts: Map<String, String>,
}

impl NumFmts {
    pub fn new() -> Self {
        NumFmts {
            builtin_numfmts: convert_args!(btreemap!(
                "0" =>  "general",
	            "1" =>  "0",
	            "2" =>  "0.00",
	            "3" =>  "#,##0",
	            "4" =>  "#,##0.00",
	            "9" =>  "0%",
	            "10" => "0.00%",
	            "11" => "0.00e+00",
	            "12" => "# ?/?",
	            "13" => "# ??/??",
	            "14" => "mm-dd-yy",
	            "15" => "d-mmm-yy",
	            "16" => "d-mmm",
	            "17" => "mmm-yy",
	            "18" => "h:mm am/pm",
	            "19" => "h:mm:ss am/pm",
	            "20" => "h:mm",
	            "21" => "h:mm:ss",
	            "22" => "m/d/yy h:mm",
	            "37" => "#,##0 ;(#,##0)",
	            "38" => "#,##0 ;[red](#,##0)",
	            "39" => "#,##0.00;(#,##0.00)",
	            "40" => "#,##0.00;[red](#,##0.00)",
	            "41" => r#"_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)"#,
	            "42" => r#"_("$"* #,##0_);_("$* \(#,##0\);_("$"* "-"_);_(@_)"#,
	            "43" => r#"_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)"#,
	            "44" => r#"_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)"#,
	            "45" => "mm:ss",
	            "46" => "[h]:mm:ss",
	            "47" => "mmss.0",
	            "48" => "##0.0e+0",
	            "49" => "@",
            )),
            defined_numfmts: Map::new(),
        }
    }

    pub fn get(&self, id: &String) -> Option<&String> {
        let mut ret = self.defined_numfmts.get(id);
        if ret.is_none() {
            ret = self.builtin_numfmts.get(id);
        }
        ret
    }

    pub fn insert(&mut self, id: &String, nfmt: &String) {
        self.defined_numfmts.insert(id.clone(), nfmt.clone());
    }
}
