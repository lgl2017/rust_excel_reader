use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_int, string_to_unsignedint};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-3.0.1
///
/// Formatting information pertaining to text alignment in cells.
///
/// Example:
/// ```
/// <xf numFmtId="0" fontId="0" applyNumberFormat="0" applyFont="1" applyFill="0" applyBorder="0" applyAlignment="1" applyProtection="0">
///     <alignment vertical="top" wrapText="1" />
/// </xf>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Alignment {
    // attributes
    /// Specifies the type of horizontal alignment in cells
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.horizontalalignmentvalues?view=openxml-3.0.1
    pub horizontal: Option<String>,

    /// Indicates the number of spaces (of the normal style font) of indentation for text in a cell.
    ///
    /// An integer value, where an increment of 1 represents 3 spaces.
    ///
    /// The number of spaces to indent is calculated as following:
    /// ```
    /// Number of spaces to indent = indent value * 3
    /// ```
    pub indent: Option<u64>,

    /// A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text.
    // tag: justifyLastLine
    pub justify_last_line: Option<bool>,

    /// An integer value indicating whether the reading order (bidirectionality) of the cell is left-to-right, right-to-left, or context dependent.
    /// 0 - Context Dependent - reading order is determined by scanning the text for the first non-whitespace character: if it is a strong right-to-left character, the reading order is right-to-left; otherwise, the reading order left-to-right.
    /// 1 - Left-to-Right- reading order is left-to-right in the cell, as in English.
    /// 2 - Right-to-Left - reading order is right-to-left in the cell, as in Hebrew.
    // tag: readingOrder
    pub reading_order: Option<u64>,

    /// An integer value (used only in a dxf element) to indicate the additional number of spaces of indentation to adjust for text in a cell.
    // tag: relativeIndent
    pub relative_indent: Option<i64>,

    /// A boolean value indicating if the displayed text in the cell should be shrunk to fit the cell width.
    /// Not applicable when a cell contains multiple lines of text.
    // tag: shrinkToFit
    pub shrink_to_fit: Option<bool>,

    /// Text rotation in cells. Expressed in degrees.
    ///
    /// For 0 - 90, the value represents degrees above horizon.
    /// For 91-180,  the value represents degrees below horizon.
    // tag: textRotation
    pub text_rotation: Option<u64>,

    /// Vertical alignment in cells.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticalalignmentvalues?view=openxml-3.0.1
    pub vertical: Option<String>,

    /// A boolean value indicating if the text in a cell should be line-wrapped within the cell.
    // tag: wrapText
    pub wrap_text: Option<bool>,
}

impl Alignment {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut alignment = Self {
            horizontal: None,
            indent: None,
            justify_last_line: None,
            reading_order: None,
            relative_indent: None,
            shrink_to_fit: None,
            text_rotation: None,
            vertical: None,
            wrap_text: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"horizontal" => alignment.horizontal = Some(string_value),
                        b"indent" => alignment.indent = string_to_unsignedint(&string_value),
                        b"justifyLastLine" => {
                            alignment.justify_last_line = string_to_bool(&string_value)
                        }
                        b"readingOrder" => {
                            alignment.reading_order = string_to_unsignedint(&string_value)
                        }
                        b"relativeIndent" => {
                            alignment.relative_indent = string_to_int(&string_value)
                        }
                        b"shrinkToFit" => alignment.shrink_to_fit = string_to_bool(&string_value),
                        b"textRotation" => {
                            alignment.text_rotation = string_to_unsignedint(&string_value)
                        }
                        b"vertical" => alignment.vertical = Some(string_value),
                        b"wrapText" => alignment.wrap_text = string_to_bool(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        return Ok(alignment);
    }
}
