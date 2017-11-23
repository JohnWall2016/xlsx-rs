use std::collections::BTreeMap as Map;

#[derive(Debug)]
pub struct WorkBook {
    pub date1904: bool,
    sheets: Vec<Option<Sheet>>,
    name_map: Map<String, usize>,
}

impl WorkBook {
    pub fn new() -> Self {
        WorkBook {
            date1904: false,
            sheets: Vec::new(),
            name_map: Map::new(),
        }
    }
}

#[derive(Debug)]
pub struct Sheet {
    
}
