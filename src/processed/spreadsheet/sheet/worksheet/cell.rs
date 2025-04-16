use crate::common_types::Coordinate;

use super::{cell_property::CellProperty, cell_value::CellValueType};

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub coordinate: Coordinate,
    pub value: CellValueType,
    pub property: CellProperty,
}
