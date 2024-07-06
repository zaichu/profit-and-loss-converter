#[derive(Clone, Debug)]
pub struct CoordinateItem {
    pub col: u32,
    pub row: u32,
}

pub trait Coordinate {
    fn new_coordinate(&self) -> CoordinateItem;
}

impl Coordinate for (u32, u32) {
    fn new_coordinate(&self) -> CoordinateItem {
        CoordinateItem {
            col: self.0,
            row: self.1,
        }
    }
}

impl Coordinate for (&u32, u32) {
    fn new_coordinate(&self) -> CoordinateItem {
        CoordinateItem {
            col: *self.0,
            row: self.1,
        }
    }
}

impl Coordinate for (u32, &u32) {
    fn new_coordinate(&self) -> CoordinateItem {
        CoordinateItem {
            col: self.0,
            row: *self.1,
        }
    }
}

impl Coordinate for (&u32, &u32) {
    fn new_coordinate(&self) -> CoordinateItem {
        CoordinateItem {
            col: *self.0,
            row: *self.1,
        }
    }
}
