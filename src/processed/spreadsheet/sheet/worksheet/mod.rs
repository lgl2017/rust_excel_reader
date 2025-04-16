pub mod calculation_reference;
pub mod cell;
pub mod cell_property;
pub mod cell_value;
pub mod table;

use anyhow::{bail, Context};
use calculation_reference::CalculationReferenceMode;
use cell::Cell;
use cell_property::CellProperty;
use cell_value::CellValueType;
use std::u64;
use table::Table;

use crate::{
    common_types::{Coordinate, Dimension},
    raw::{
        drawing::scheme::color_scheme::ColorScheme,
        spreadsheet::{
            shared_string::shared_string_item::SharedStringItem,
            sheet::worksheet::{
                cell::Cell as RawCell, column_information::ColumnInformation as RawCol,
                row::Row as RawRow, Worksheet as RawWorksheet,
            },
            stylesheet::{
                format::{alignment::Alignment, cell_format::CellFormat, protection::Protection},
                StyleSheet,
            },
            table::Table as RawTable,
        },
    },
};

#[derive(Debug, Clone, PartialEq)]
pub struct Worksheet {
    pub name: String,
    pub sheet_id: u64,

    pub dimension: Option<Dimension>,
    pub merged_cells: Vec<Dimension>,

    pub tables: Vec<Table>,

    /// Value that indicates whether to use a 1900 or 1904 date base when converting serial values in the workbook to dates. [Note: If the dateCompatibility attribute is 0 or false, this attribute is ignored. end note]
    ///
    /// - true: workbook uses the 1904 backward compatibility date system.
    /// - false:  the workbook uses a date system based in 1900, as specified by the value of the dateCompatibility attribute.
    /// The default value for this attribute is false.
    pub is_1904: bool,

    /// Calculation Reference Mode
    pub calculation_reference_mode: Option<CalculationReferenceMode>,

    // private
    raw_sheet: Box<RawWorksheet>,
    shared_string_items: Box<Vec<SharedStringItem>>,
    stylesheet: Box<StyleSheet>,
    color_scheme: Box<Option<ColorScheme>>,
}

impl Worksheet {
    pub(crate) fn from_raw(
        name: String,
        sheet_id: u64,
        worksheet: RawWorksheet,
        tables: Vec<RawTable>,
        is_1904: bool,
        calculation_reference_mode: Option<CalculationReferenceMode>,
        shared_string_items: Vec<SharedStringItem>,
        stylesheet: StyleSheet,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let default_table_style_name = if let Some(style) = stylesheet.clone().table_styles {
            style.default_table_style
        } else {
            None
        };

        let tables: Vec<Table> = tables
            .into_iter()
            .map(|t| Table::from_raw(t, default_table_style_name.clone()))
            .collect();

        return Self {
            name,
            sheet_id,
            dimension: worksheet.dimension,
            merged_cells: worksheet.merge_cells.clone().unwrap_or(vec![]),
            tables,
            is_1904,
            calculation_reference_mode,
            raw_sheet: Box::new(worksheet),
            shared_string_items: Box::new(shared_string_items),
            stylesheet: Box::new(stylesheet),
            color_scheme: Box::new(color_scheme),
        };
    }
}

impl Worksheet {
    /// get cell value and styles for a specific coordinate.
    ///
    /// The style here ignoring table settings.
    /// If the cell is within a table that has different header row colors, column/row stripes, and etc.
    /// The appearance can be different.
    pub fn get_cell(&mut self, coordinate: Coordinate) -> anyhow::Result<Cell> {
        let row = self.get_raw_row(coordinate)?;
        let cell = self.get_raw_cell(coordinate, row.clone())?;
        let col = self.get_raw_col_info(coordinate);

        let cell_value = CellValueType::from_raw(
            cell.clone(),
            *self.shared_string_items.clone(),
            *self.stylesheet.clone(),
            *self.color_scheme.clone(),
        )?;

        let num_format_id = self.get_id(cell.clone(), row.clone(), col.clone(), &|x| {
            self.get_number_format_id_helper(x)
        });

        let fill_id = self.get_id(cell.clone(), row.clone(), col.clone(), &|x| {
            self.get_fill_id_helper(x)
        });
        let border_id = self.get_id(cell.clone(), row.clone(), col.clone(), &|x| {
            self.get_border_id_helper(x)
        });
        let font_id = self.get_id(cell.clone(), row.clone(), col.clone(), &|x| {
            self.get_font_id_helper(x)
        });
        let alignment = self.get_alignment(cell.clone(), row.clone(), col.clone());
        let protection = self.get_protection(cell.clone(), row.clone(), col.clone());

        let cell_property = CellProperty::from_raw(
            cell.clone(),
            row.clone(),
            col.clone(),
            fill_id,
            font_id,
            border_id,
            num_format_id,
            alignment,
            protection,
            (*self.raw_sheet).sheet_format_properties.clone(),
            *self.stylesheet.clone(),
            *self.color_scheme.clone(),
        );

        println!("\ncell: {:?}", coordinate);
        println!("num_format_id: {:?}", num_format_id);
        println!("font_id: {:?}", font_id);
        println!("fill_id: {:?}", fill_id);
        println!("border_id: {:?}", border_id);
        println!(" \ncell value {:?}.", cell_value);

        println!(" \ncell_property ");
        println!(
            "(width, height) : {:?}",
            (cell_property.width, cell_property.height)
        );
        println!("hidden : {:?}", cell_property.hidden);
        println!("show_phonetic : {:?}", cell_property.show_phonetic);
        println!("font : {:?}", cell_property.font);
        println!("border : {:?}", cell_property.border);
        println!("fill : {:?}", cell_property.fill);
        println!("alignment : {:?}", cell_property.alignment);

        Ok(Cell {
            value: cell_value,
            property: cell_property,
        })
    }
}

impl Worksheet {
    /// get cell alignment information
    fn get_protection(
        &self,
        cell: RawCell,
        row_info: RawRow,
        col_info: Option<RawCol>,
    ) -> Option<Protection> {
        if let Some(n) = cell.style {
            if let Some(protection) = self.get_protection_helper(n) {
                return Some(protection);
            }
        }

        if let Some(n) = row_info.style {
            if let Some(protection) = self.get_protection_helper(n) {
                return Some(protection);
            }
        }

        if let Some(col) = col_info {
            if let Some(n) = col.style {
                if let Some(protection) = self.get_protection_helper(n) {
                    return Some(protection);
                }
            }
        }

        return None;
    }

    /// get protection for a cellXfs' xf_id.
    ///
    /// None if not specified or applyFont is set to false
    fn get_protection_helper(&self, xf_id: u64) -> Option<Protection> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.protection.is_some() && cell_format.apply_protection == Some(true) {
            return cell_format.protection;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.protection.is_some() && cell_format.apply_protection.is_none() {
                return cell_format.protection;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.protection.is_some()
            && cell_style_format.apply_protection.unwrap_or(true) == true
        {
            return cell_style_format.protection;
        }

        return None;
    }

    /// get cell alignment information
    fn get_alignment(
        &self,
        cell: RawCell,
        row_info: RawRow,
        col_info: Option<RawCol>,
    ) -> Option<Alignment> {
        if let Some(n) = cell.style {
            if let Some(alignment) = self.get_alignment_helper(n) {
                return Some(alignment);
            }
        }

        if let Some(n) = row_info.style {
            if let Some(alignment) = self.get_alignment_helper(n) {
                return Some(alignment);
            }
        }

        if let Some(col) = col_info {
            if let Some(n) = col.style {
                if let Some(alignment) = self.get_alignment_helper(n) {
                    return Some(alignment);
                }
            }
        }

        return None;
    }

    /// get alignment for a cellXfs' xf_id.
    ///
    /// None if not specified or applyFont is set to false
    fn get_alignment_helper(&self, xf_id: u64) -> Option<Alignment> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.alignment.is_some() && cell_format.apply_alignment == Some(true) {
            return cell_format.alignment;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.alignment.is_some() && cell_format.apply_alignment.is_none() {
                return cell_format.alignment;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.alignment.is_some()
            && cell_style_format.apply_alignment.unwrap_or(true) == true
        {
            return cell_style_format.alignment;
        }

        return None;
    }

    /// get id for a cell
    ///
    /// helper function:
    /// * `get_fill_id_helper`
    /// * `get_border_id_helper`
    /// * `get_font_id_helper`
    /// * `get_number_format_id_helper`
    fn get_id(
        &self,
        cell: RawCell,
        row_info: RawRow,
        col_info: Option<RawCol>,
        helper_function: &dyn Fn(u64) -> Option<u64>,
    ) -> Option<u64> {
        if let Some(n) = cell.style {
            if let Some(id) = helper_function(n) {
                return Some(id);
            }
        }

        if let Some(n) = row_info.style {
            if let Some(id) = helper_function(n) {
                return Some(id);
            }
        }

        if let Some(col) = col_info {
            if let Some(n) = col.style {
                if let Some(id) = helper_function(n) {
                    return Some(id);
                }
            }
        }

        return None;
    }

    /// get border id for a cellXfs' xf_id.
    ///
    /// None if not specified or applyFont is set to false
    fn get_border_id_helper(&self, xf_id: u64) -> Option<u64> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.border_id.is_some() && cell_format.apply_border == Some(true) {
            return cell_format.border_id;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.border_id.is_some() && cell_format.apply_border.is_none() {
                return cell_format.border_id;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.border_id.is_some()
            && cell_style_format.apply_border.unwrap_or(true) == true
        {
            return cell_style_format.border_id;
        }

        return None;
    }

    /// get fill id for a cellXfs' xf_id.
    ///
    /// None if not specified or applyFill is set to false
    fn get_fill_id_helper(&self, xf_id: u64) -> Option<u64> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.fill_id.is_some() && cell_format.apply_fill == Some(true) {
            return cell_format.fill_id;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.fill_id.is_some() && cell_format.apply_fill.is_none() {
                return cell_format.fill_id;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.fill_id.is_some()
            && cell_style_format.apply_fill.unwrap_or(true) == true
        {
            return cell_style_format.fill_id;
        }

        return None;
    }

    /// get font id for a cellXfs' xf_id.
    ///
    /// None if not specified or applyFont is set to false
    fn get_font_id_helper(&self, xf_id: u64) -> Option<u64> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.font_id.is_some() && cell_format.apply_font == Some(true) {
            return cell_format.font_id;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.font_id.is_some() && cell_format.apply_font.is_none() {
                return cell_format.font_id;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.font_id.is_some()
            && cell_style_format.apply_font.unwrap_or(true) == true
        {
            return cell_style_format.font_id;
        }

        return None;
    }

    /// get number format id for a cellXfs' xf_id.
    ///
    /// None if not specified or apply_number_format is set to false
    fn get_number_format_id_helper(&self, xf_id: u64) -> Option<u64> {
        let Some(cell_format) = self.get_cell_format(xf_id) else {
            return None;
        };
        if cell_format.num_fmt_id.is_some() && cell_format.apply_number_format == Some(true) {
            return cell_format.num_fmt_id;
        }

        // check if there is any reference to cellStyleXfs
        let Some(cell_style_format_id) = cell_format.xf_id else {
            if cell_format.num_fmt_id.is_some() && cell_format.apply_number_format.is_none() {
                return cell_format.num_fmt_id;
            } else {
                return None;
            }
        };

        let Some(cell_style_format) = self.get_cell_style_format(cell_style_format_id) else {
            return None;
        };

        if cell_style_format.num_fmt_id.is_some()
            && cell_style_format.apply_number_format.unwrap_or(true) == true
        {
            return cell_style_format.num_fmt_id;
        }

        return None;
    }

    fn get_raw_cell(&mut self, coordinate: Coordinate, row: RawRow) -> anyhow::Result<RawCell> {
        let cells = row.cells.context("cells not availble in row.")?;
        let col_coordintae = coordinate
            .col
            .checked_sub(1)
            .context("col coordinate out of range.")?;
        let col_coordintae = TryInto::<usize>::try_into(col_coordintae)?;

        if col_coordintae >= cells.len() {
            bail!("row coordinate out of range.")
        }
        let raw_cell = cells[col_coordintae].clone();
        if raw_cell.coordinate != Some(coordinate) {
            bail!("inconsistent coordinate.")
        }
        return Ok(raw_cell);
    }

    fn get_raw_col_info(&mut self, coordinate: Coordinate) -> Option<RawCol> {
        let cols = self.raw_sheet.column_infos.clone().unwrap_or(vec![]);

        let col_coordintae = coordinate.col;

        for col in cols {
            let min = col.min_column.unwrap_or(u64::MIN);
            let max = col.max_column.unwrap_or(u64::MAX);
            if (min..=max).contains(&col_coordintae) {
                return Some(col);
            }
        }

        return None;
    }

    fn get_raw_row(&mut self, coordinate: Coordinate) -> anyhow::Result<RawRow> {
        let sheet_data = self
            .raw_sheet
            .clone()
            .sheet_data
            .context("Sheet data does not exists.")?;

        let rows = sheet_data.rows.unwrap_or(vec![]);
        let row_coordinate = coordinate
            .row
            .checked_sub(1)
            .context("row coordinate out of range.")?;

        let row_coordinate = TryInto::<usize>::try_into(row_coordinate)?;

        if row_coordinate >= rows.len() {
            bail!("row coordinate out of range.")
        }
        let row = rows[row_coordinate].clone();
        return Ok(row);
    }

    fn get_cell_format(&self, xf_id: u64) -> Option<CellFormat> {
        let Ok(xf_id) = TryInto::<usize>::try_into(xf_id) else {
            return None;
        };

        return self.stylesheet.get_cell_format(xf_id);
    }

    fn get_cell_style_format(&self, xf_id: u64) -> Option<CellFormat> {
        let Ok(cell_style_format_id) = TryInto::<usize>::try_into(xf_id) else {
            return None;
        };

        return self.stylesheet.get_cell_style_format(cell_style_format_id);
    }
}
