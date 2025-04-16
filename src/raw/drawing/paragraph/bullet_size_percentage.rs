use quick_xml::events::BytesStart;

use crate::helper::{extract_val_attribute, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizepercentage?view=openxml-3.0.1
///
/// This element specifies the size in percentage of the surrounding text to be used on bullet characters within a given paragraph.
/// The size is specified using a percentage where 1000 is equal to 1 percent of the font size and 100000 is equal to 100 percent font of the font size.
/// Example:
/// ```
/// <a:pPr …>
///      <a:buSzPct val="111000"/>
/// </a:pPr>
/// ```
// tag: buSzPct
#[derive(Debug, Clone, PartialEq)]
pub struct BulletSizePercentage {
    // Attributes
    /// Specifies the percentage of the text size that this bullet should be. It is specified here in terms of 100% being equal to 100000 and 1% being specified in increments of 1000.
    /// This attribute should not be lower than 25%, or 25000 and not be higher than 400%, or 400000.
    // val (Value)
    pub val: Option<i64>,
}

impl BulletSizePercentage {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut bullet = Self { val: None };

        if let Some(val) = extract_val_attribute(e)? {
            bullet.val = string_to_int(&val);
        }

        Ok(bullet)
    }
}
