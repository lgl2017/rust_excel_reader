use quick_xml::events::BytesStart;

use crate::{
    helper::{extract_val_attribute, string_to_unsignedint},
    raw::drawing::st_types::STPositivePercentage,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spacingpercent?view=openxml-3.0.1
///
/// This element specifies the amount of white space that is to be used between lines and paragraphs in the form of a percentage of the text size.
///
/// Example:
/// ```
/// <a:pPr>
///     <a:lnSpc>
///         <a:spcPct val="200%"/>
///     </a:lnSpc>
/// </a:pPr>
/// ```
// tag: spcPct
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSpacingPercent {
    // Attributes
    /// Specifies the percentage of the size that the white space should be.
    // val (Value)
    pub val: Option<STPositivePercentage>,
}

impl XlsxSpacingPercent {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut spacing = Self { val: None };

        if let Some(val) = extract_val_attribute(e)? {
            spacing.val = string_to_unsignedint(&val);
        }

        Ok(spacing)
    }
}
