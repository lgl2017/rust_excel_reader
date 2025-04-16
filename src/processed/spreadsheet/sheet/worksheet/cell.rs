use super::{cell_property::CellProperty, cell_value::CellValueType};

#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    pub value: CellValueType,
    pub property: CellProperty,
}
