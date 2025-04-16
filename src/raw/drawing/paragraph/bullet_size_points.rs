use quick_xml::events::BytesStart;

use crate::helper::{extract_val_attribute, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizepoints?view=openxml-3.0.1
///
/// This element specifies the size in points to be used on bullet characters within a given paragraph.
/// The size is specified using the points where 100 is equal to 1 point font and 1200 is equal to 12 point font.
///
/// Example:
/// ```
/// <a:pPr …>
///      <a:buSzPts val="1400"/>
/// </a:pPr>
/// ```
// tag: buSzPts
#[derive(Debug, Clone, PartialEq)]
pub struct BulletSizePoints {
    // Attributes
    /// Specifies the size of the bullets in point size.
    ///  Whole points are specified in increments of 100 starting with 100 being a point size of 1.
    // val (Value)
    pub val: Option<i64>,
}

impl BulletSizePoints {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let mut bullet = Self { val: None };
        if let Some(val) = extract_val_attribute(e)? {
            bullet.val = string_to_int(&val);
        }
        Ok(bullet)
    }
}
