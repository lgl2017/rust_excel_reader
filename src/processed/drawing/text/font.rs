use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::text::{
    default_text_run_properties::XlsxTextRunProperties,
    font::{base_font::XlsxBaseFont, XlsxFontBase},
};

use super::text_run_properties::TextRunProperties;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
///
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Font {
    /// LatinFont typeface
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.latinfont?view=openxml-3.0.1
    pub latin: String,

    /// ComplexScriptFont typeface
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.complexscriptfont?view=openxml-3.0.1
    pub complex_script: String,

    /// EastAsianFont typeface
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.eastasianfont?view=openxml-3.0.1
    pub east_asian: String,

    /// supplemental fonts (script, typeface)
    #[cfg_attr(feature = "serde", serde(skip_serializing))]
    pub supplemental_fonts: BTreeMap<String, String>,
}

impl Font {
    pub(crate) fn default() -> Self {
        return Self {
            latin: "Aptos Display".to_string(),
            complex_script: "".to_string(),
            east_asian: "".to_string(),
            supplemental_fonts: BTreeMap::new(),
        };
    }

    pub(crate) fn from_raw(raw: Option<XlsxFontBase>) -> Option<Self> {
        let Some(raw) = raw else {
            return None;
        };

        let mut supplemental: BTreeMap<String, String> = BTreeMap::new();

        for font in raw.clone().font.unwrap_or(vec![]).into_iter() {
            let (Some(script), Some(tf)) = (font.script, font.typeface) else {
                continue;
            };
            supplemental.insert(script, tf);
        }

        return Some(Self {
            latin: Self::base_font_to_string(raw.clone().latin).unwrap_or("Arial".to_string()),
            complex_script: Self::base_font_to_string(raw.clone().cs).unwrap_or(String::new()),
            east_asian: Self::base_font_to_string(raw.clone().ea).unwrap_or(String::new()),
            supplemental_fonts: supplemental,
        });
    }

    pub(crate) fn from_text_run_properties(
        raw: XlsxTextRunProperties,
        default_properties: Option<TextRunProperties>,
        ref_font: Option<Font>,
    ) -> Self {
        let reference = if let Some(r) = ref_font {
            r
        } else {
            if let Some(default_properties) = default_properties {
                default_properties.font_typeface
            } else {
                Self::default()
            }
        };

        return Self {
            latin: Self::base_font_to_string(raw.clone().latin).unwrap_or(reference.latin),
            complex_script: Self::base_font_to_string(raw.clone().cs)
                .unwrap_or(reference.complex_script),
            east_asian: Self::base_font_to_string(raw.clone().ea).unwrap_or(reference.east_asian),
            supplemental_fonts: reference.supplemental_fonts,
        };
    }

    fn base_font_to_string(font: Option<XlsxBaseFont>) -> Option<String> {
        let Some(font) = font else {
            return None;
        };
        return font.typeface;
    }
}
