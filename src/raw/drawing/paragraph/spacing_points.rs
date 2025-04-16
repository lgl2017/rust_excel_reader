use quick_xml::events::BytesStart;

use crate::helper::{extract_val_attribute, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpoints?view=openxml-3.0.1
///
/// This element specifies the amount of white space that is to be used between lines and paragraphs in the form of a text point size.
///
/// Example:
/// ```
/// <a:pPr>
///     <a:lnSpc>
///         <a:spcPts val="1400"/>
///     </a:lnSpc>
/// </a:pPr>
/// ```
// tag: spcPts
#[derive(Debug, Clone, PartialEq)]
pub struct SpacingPoints {
    // Attributes
    /// Specifies the size of the white space in point size
    // val (Value)
    pub val: Option<i64>,
}

impl SpacingPoints {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut spacing = Self { val: None };

        if let Some(val) = extract_val_attribute(e)? {
            spacing.val = string_to_int(&val);
        }

        Ok(spacing)
    }
}
