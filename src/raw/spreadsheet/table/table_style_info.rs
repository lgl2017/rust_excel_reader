use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::string_to_bool;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyleinfo?view=openxml-3.0.1
///
/// This element describes which style is used to display this table, and specifies which portions of the table have the style applied.
///
/// Styles define a set of formatting properties that can be easily referenced by cells or other objects in the spreadsheet application.
/// A style can be applied to a table, but tables can define specific parts of the table that should not have the style applied independently of other table parts.
/// For instance a table can not apply the row striping of the style, and can not show the style's formatting of the last column, but will apply the column striping and the formatting to the first column.
///
/// Example:
/// ```
/// <tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />
/// ```
///
/// tableStyleInfo (Table Style)
#[derive(Debug, Clone, PartialEq)]
pub struct TableStyleInfo {
    // Attributes
    /// name (Style Name)
    ///
    /// A string representing the name of the table style to use with this table.
    /// If the style name does not correspond to the name of a table style then the spreadsheet application should use default style.
    pub name: Option<String>,

    /// showColumnStripes (Show Column Stripes)
    ///
    /// A Boolean indicating whether column stripe formatting is applied.
    /// True when style column stripe formatting is applied, false otherwise.
    pub show_column_stripes: Option<bool>,

    /// showFirstColumn (Show First Column)
    ///
    /// A Boolean indicating whether the first column in the table should have the style applied.
    /// True if the first column has the style applied, false otherwise.
    pub show_first_column: Option<bool>,

    /// showLastColumn (Show Last Column)
    ///
    /// A Boolean indicating whether the last column in the table should have the style applied.
    /// True if the last column has the style applied, false otherwise.
    pub show_last_column: Option<bool>,

    /// showRowStripes (Show Row Stripes)
    ///
    /// A Boolean indicating whether row stripe formatting is applied.
    /// True when style row stripe formatting is applied, false otherwise.
    pub show_row_stripes: Option<bool>,
}

impl TableStyleInfo {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut info = Self {
            name: None,
            show_column_stripes: None,
            show_first_column: None,
            show_last_column: None,
            show_row_stripes: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => info.name = Some(string_value),
                        b"showColumnStripes" => {
                            info.show_column_stripes = string_to_bool(&string_value);
                        }
                        b"showFirstColumn" => {
                            info.show_first_column = string_to_bool(&string_value);
                        }
                        b"showLastColumn" => {
                            info.show_last_column = string_to_bool(&string_value);
                        }
                        b"showRowStripes" => {
                            info.show_row_stripes = string_to_bool(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(info)
    }
}
