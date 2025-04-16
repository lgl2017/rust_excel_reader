use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_float, string_to_unsignedint};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetformatproperties?view=openxml-3.0.1
///
/// Sheet formatting properties.
///
/// Example:
/// ```
/// <sheetFormatPr defaultColWidth="16.3333" defaultRowHeight="19.9" customHeight="1" outlineLevelRow="0" outlineLevelCol="0" />
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct SheetFormatProperties {
    // Attributes
    /// baseColWidth (Base Column Width)
    ///
    /// Specifies the number of characters of the maximum digit width of the normal style's font.
    /// This value does not include margin padding or extra padding for gridlines.
    /// It is only the number of characters.
    pub base_col_width: Option<u64>,

    /// customHeight (Custom Height)
    ///
    /// 'True' if defaultRowHeight value has been manually set, or is different from the default value.
    pub custom_height: Option<bool>,

    /// defaultColWidth (Default Column Width)
    ///
    /// Default column width measured as the number of characters of the maximum digit width of the normal style's font.
    ///
    /// If the user has not set this manually, then it can be calculated:
    /// ```
    /// defaultColWidth = baseColumnWidth + {margin padding (2 pixels on each side, totalling 4 pixels)} + {gridline (1pixel)}
    /// ```
    ///
    /// If the user has set this manually, then there is no calculation, and simply a value is specified.
    pub default_col_width: Option<f64>,

    /// defaultRowHeight (Default Row Height)
    ///
    /// Default row height measured in point size.
    /// Optimization so we don't have to write the height on all rows.
    ///
    /// This can be written out if most rows have custom height, to achieve the optimization.
    ///
    /// When the row height of all rows in a sheet is the default value, then that value is written here, and customHeight is not set.
    /// If a few rows have a different height, that information is written directly on each row.
    ///
    /// However, if most or all of the rows in the sheet have the same height, but that height isn't the default height, then that height value should be written here (as an optimization), and the customHeight flag should also be set.
    /// In this case, all rows having this height do not need to express the height, only rows whose height differs from this value need to be explicitly expressed.
    pub default_row_height: Option<f64>,

    /// outlineLevelCol (Column Outline Level)
    ///
    /// Highest number of outline levels for columns in this sheet.
    /// These values shall be in synch with the actual sheet outline levels.
    ///
    /// unsignedByte
    pub outline_level_col: Option<u64>,

    /// outlineLevelRow (Maximum Outline Row)
    ///
    /// Highest number of outline level for rows in this sheet.
    /// These values shall be in synch with the actual sheet outline levels.
    ///
    /// unsignedByte
    pub outline_level_row: Option<u64>,

    /// thickBottom (Thick Bottom Border)
    ///
    /// 'True' if rows have a thick bottom border by default.
    pub thick_bottom: Option<bool>,

    /// thickTop (Thick Top Border)
    ///
    /// 'True' if rows have a thick top border by default.
    pub thick_top: Option<bool>,

    /// zeroHeight (Hidden By Default)
    ///
    /// 'True' if rows are hidden by default.
    /// This setting is an optimization used when most rows of the sheet are hidden.
    /// In this case, instead of writing out every row and specifying hidden, it is much shorter to only write out the rows that are not hidden, and specify here that rows are hidden by default, and only not hidden if specified.
    pub zero_height: Option<bool>,
}

impl SheetFormatProperties {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut properties = Self {
            base_col_width: None,
            custom_height: None,
            default_col_width: None,
            default_row_height: None,
            outline_level_col: None,
            outline_level_row: None,
            thick_bottom: None,
            thick_top: None,
            zero_height: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"baseColWidth" => {
                            properties.base_col_width = string_to_unsignedint(&string_value);
                        }
                        b"customHeight" => {
                            properties.custom_height = string_to_bool(&string_value);
                        }
                        b"defaultColWidth" => {
                            properties.default_col_width = string_to_float(&string_value);
                        }
                        b"defaultRowHeight" => {
                            properties.default_row_height = string_to_float(&string_value);
                        }
                        b"outlineLevelCol" => {
                            properties.outline_level_col = string_to_unsignedint(&string_value);
                        }
                        b"outlineLevelRow" => {
                            properties.outline_level_row = string_to_unsignedint(&string_value);
                        }
                        b"thickBottom" => {
                            properties.thick_bottom = string_to_bool(&string_value);
                        }
                        b"thickTop" => {
                            properties.thick_top = string_to_bool(&string_value);
                        }
                        b"zeroHeight" => {
                            properties.zero_height = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(properties)
    }
}
