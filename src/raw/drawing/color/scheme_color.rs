use quick_xml::events::BytesStart;
use std::io::Read;

use crate::common_types::HexColor;
use crate::excel::XmlReader;

use crate::helper::{extract_val_attribute, hex_to_rgba, rgba_to_hex};
use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;

use super::color_transforms::{apply_color_transformations, XlsxColorTransform};

/// SchemeColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolor?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:schemeClr val="phClr">
///     <a:lumMod val="110000" />
///     <a:satMod val="105000" />
///     <a:tint val="67000" />
/// </a:schemeClr>
/// ```
///
/// tag: schemeClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSchemeColor {
    // attributes
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolorvalues?view=openxml-3.0.1
    pub val: Option<String>,

    // children
    pub color_transforms: Option<Vec<XlsxColorTransform>>,
}

impl XlsxSchemeColor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let val = extract_val_attribute(e)?;
        let mut color = Self {
            val,
            color_transforms: None,
        };
        color.color_transforms = Some(XlsxColorTransform::load_list(reader, b"schemeClr")?);

        Ok(color)
    }
}

impl XlsxSchemeColor {
    /// SchemeColorValues: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.schemecolorvalues?view=openxml-3.0.1
    pub(crate) fn to_hex(
        &self,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<HexColor> {
        let Some(val) = self.val.clone() else {
            return None;
        };

        if &val == "phClr" {
            let Some(hex) = ref_color else {
                return None;
            };
            return self.to_hex_helper(hex);
        }

        let Some(scheme) = color_scheme else {
            return None;
        };

        let Some(hex) = (match val.as_ref() {
            "lt1" => scheme.lt1.clone(),
            "dk1" => scheme.dk1.clone(),
            "lt2" => scheme.lt2.clone(),
            "dk2" => scheme.dk2.clone(),
            "accent1" => scheme.accent1.clone(),
            "accent2" => scheme.accent2.clone(),
            "accent3" => scheme.accent3.clone(),
            "accent4" => scheme.accent4.clone(),
            "accent5" => scheme.accent5.clone(),
            "accent6" => scheme.accent6.clone(),
            "hlink" => scheme.hlink.clone(),
            "folHlink" => scheme.fol_hlink.clone(),
            "tx1" => Some("000000ff".to_string()),
            "tx2" => Some("0e2841ff".to_string()),
            "bg1" => Some("ffffffff".to_string()),
            "bg2" => Some("e8e8e8ff".to_string()),
            "phClr" => ref_color,
            _ => None,
        }) else {
            return None;
        };

        return self.to_hex_helper(hex);
    }

    fn to_hex_helper(&self, hex: HexColor) -> Option<HexColor> {
        let mut hex = hex;

        if hex.len() == 6 {
            hex = format!("{}ff", hex);
        };

        let Ok(mut rgba) = hex_to_rgba(&hex, Some(false)) else {
            return Some(hex);
        };

        rgba = apply_color_transformations(rgba, self.color_transforms.clone().unwrap_or(vec![]));

        match rgba_to_hex(rgba, Some(false)) {
            Ok(hex) => return Some(hex),
            Err(_) => return Some(hex),
        }
    }
}
