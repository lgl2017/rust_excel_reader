#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use crate::raw::spreadsheet::table::table_style_info::XlsxTableStyleInfo;

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct TableStyle {
    /// A string representing the name of the table style to use with this table.
    ///
    /// If the style name does not correspond to the name of a table style then the spreadsheet application should use default style.
    pub name: Option<String>,

    /// A Boolean indicating whether column stripe formatting is applied.
    /// True when style column stripe formatting is applied, false otherwise.
    pub show_column_stripes: bool,

    /// A Boolean indicating whether the first column in the table should have the style applied.
    /// True if the first column has the style applied, false otherwise.
    pub style_first_column: bool,

    /// A Boolean indicating whether the last column in the table should have the style applied.
    /// True if the last column has the style applied, false otherwise.
    pub style_last_column: bool,

    /// showRowStripes (Show Row Stripes)
    ///
    /// A Boolean indicating whether row stripe formatting is applied.
    /// True when style row stripe formatting is applied, false otherwise.
    pub show_row_stripes: bool,
}

impl TableStyle {
    pub(crate) fn from_raw(
        style: Option<XlsxTableStyleInfo>,
        default_table_style: Option<String>,
    ) -> Self {
        let Some(style) = style else {
            return Self {
                name: None,
                show_column_stripes: false,
                style_first_column: false,
                style_last_column: false,
                show_row_stripes: false,
            };
        };
        let name = if let Some(n) = style.name {
            Some(n)
        } else {
            default_table_style
        };
        return Self {
            name,
            show_column_stripes: style.show_column_stripes.unwrap_or(false),
            style_first_column: style.show_first_column.unwrap_or(false),
            style_last_column: style.show_last_column.unwrap_or(false),
            show_row_stripes: style.show_row_stripes.unwrap_or(false),
        };
    }
}
