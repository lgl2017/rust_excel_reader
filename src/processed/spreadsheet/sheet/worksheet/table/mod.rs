use table_style::TableStyle;

use crate::{common_types::Dimension, raw::spreadsheet::table::Table as RawTable};

pub mod table_style;

#[derive(Debug, Clone, PartialEq)]
pub struct Table {
    /// name of the table.
    ///
    /// This is the name that shall be used in formula references, and displayed in the UI to the spreadsheet user.
    pub display_name: String,

    /// table id
    ///
    /// Ids can be used to refer to the specific table in the workbook.
    pub table_id: u64,

    /// table dimension
    pub dimension: Dimension,

    /// List of column names
    pub columns: Vec<String>,

    /// header row count
    pub header_row_count: u64,

    /// the number of `totals rows` that is shown at the bottom of the table
    pub totals_row_count: u64,

    /// table style
    pub table_style: TableStyle,

    // private
    raw_table: Box<RawTable>,
}

impl Table {
    pub(crate) fn from_raw(table: RawTable, default_table_style: Option<String>) -> Self {
        let column_names: Vec<String> = table
            .clone()
            .table_columns
            .unwrap_or(vec![])
            .into_iter()
            .map(|c| c.name.unwrap_or("".to_string()))
            .collect();

        return Self {
            display_name: table.clone().display_name.unwrap_or("".to_string()),
            table_id: table.clone().id.unwrap_or(1),
            dimension: table.clone().r#ref.unwrap_or(Dimension::default()),
            columns: column_names,
            header_row_count: table.clone().header_row_count.unwrap_or(1),
            totals_row_count: table.clone().totals_row_count.unwrap_or(1),
            table_style: TableStyle::from_raw(table.clone().table_style_info, default_table_style),
            raw_table: Box::new(table),
        };
    }
}
