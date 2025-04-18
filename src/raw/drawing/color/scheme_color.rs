use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::helper::extract_val_attribute;

use super::color_transforms::XlsxColorTransform;

/// SchemeColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolor?view=openxml-3.0.1
/// Example:
/// ```
/// <a:schemeClr val="phClr">
///     <a:lumMod val="110000" />
///     <a:satMod val="105000" />
///     <a:tint val="67000" />
/// </a:schemeClr>
/// ```
// tag: schemeClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSchemeColor {
    // attributes
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolorvalues?view=openxml-3.0.1
    pub val: Option<String>,
    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxSchemeColor {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let val = extract_val_attribute(e)?;
        let mut color = Self {
            val,
            color_transforms: None,
        };
        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"schemeClr")?);

        Ok(color)
    }
}
