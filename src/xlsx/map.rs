use std::{fmt::Debug, collections::btree_map::{BTreeMap, Iter, IterMut}};
use std::ops::{Index, IndexMut};

pub(crate) struct IndexMap<T> (BTreeMap<usize, T>);

impl<T> IndexMap<T> {
    pub fn new() -> IndexMap<T> {
        IndexMap(BTreeMap::new())
    }

    /// Put the `value` in the `index` and return the original value.
    pub fn put(&mut self, index: usize, value: T) -> Option<T> {
        self.0.insert(index, value)
    }

    pub fn get(&self, index: usize) -> Option<&T> {
        self.0.get(&index)
    }

    pub fn get_mut(&mut self, index: usize) -> Option<&mut T> {
        self.0.get_mut(&index)
    }

    /// Insert the `value` in the `index`, the original value and the after move to next indexes.
    pub fn insert(&mut self, index: usize, value: T) {
        let mut val = Some(value);
        let mut idx = index;

        loop {
            if let Some(v) = val {
                val = self.0.insert(idx, v);
                idx += 1;
            } else {
                break;
            }
        }
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    pub fn iter(&self) -> Iter<'_, usize, T> {
        self.0.iter()
    }

    pub fn iter_mut(&mut self) -> IterMut<'_, usize, T> {
        self.0.iter_mut()
    }
}

impl<T: Debug> Debug for IndexMap<T> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_map().entries(self.iter()).finish()
    }
}

impl<T> Index<usize> for IndexMap<T> {
    type Output = T;

    #[inline]
    fn index(&self, index: usize) -> &T {
        self.get(index).expect("no entry found for key")
    }
}

impl<T> IndexMut<usize> for IndexMap<T> {
    #[inline]
    fn index_mut(&mut self, index: usize) -> &mut T {
        self.get_mut(index).expect("no entry found for key")
    }
}

#[test]
fn test_index_map() {
    let mut map = IndexMap::new();

    map.put(1, "One");
    map.put(2, "Two");
    map.put(3, "Three");
    println!("{:?}", map);

    map.insert(3, "new Three");
    println!("{:?}", map);

    map.put(4, "Four");
    println!("{:?}", map);
}