use std::collections::BTreeMap as Map;

#[derive(Debug)]
pub struct Strings {
    values: Vec<String>,
    index_map: Map<String, usize>,
}

impl Strings {
    pub fn new() -> Self {
        Strings{
            values: Vec::new(),
            index_map: Map::new(),
        }
    }

    pub fn add(self: &mut Self, str: &String) -> usize {
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

    pub fn index(self: &Self, index: usize) -> Option<&String> {
        if index >= self.values.len() {
            None
        } else {
            Some(&self.values[index])
        }
    }

    pub fn get_index(self: &Self, str: &String) -> Option<&usize> {
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

    pub fn insert(self: &mut Self, name: String, rgb_color: String) -> Option<String> {
        self.values.insert(name, rgb_color)
    }

    pub fn get(self: &Self, name: &String) -> Option<&String> {
        self.values.get(name)
    }
}
