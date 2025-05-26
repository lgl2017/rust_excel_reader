#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::{HexColor, Text},
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::{scheme::color_scheme::XlsxColorScheme, text::text_run::XlsxTextRun},
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{font::Font, text_run_properties::TextRunProperties};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.run?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:r>
///     <a:rPr kumimoji="0" lang="en-US" sz="1100" b="0" i="0" u="none"
///         strike="noStrike" cap="none" spc="0" normalizeH="0" baseline="0">
///         <a:ln>
///             <a:noFill />
///         </a:ln>
///         <a:solidFill>
///             <a:srgbClr val="000000" />
///         </a:solidFill>
///         <a:effectLst />
///         <a:uFillTx />
///         <a:latin typeface="+mn-lt" />
///         <a:ea typeface="+mn-ea" />
///         <a:cs typeface="+mn-cs" />
///         <a:sym typeface="Helvetica Neue" />
///     </a:rPr>
///     <a:t>Text</a:t>
/// </a:r>
/// ```
///
/// r (Text Run)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct TextRun {
    // rPr (Text Run Properties)
    pub run_properties: TextRunProperties,

    // t (Text String)
    pub text: Text,
}

impl TextRun {
    pub(crate) fn from_raw(
        raw: XlsxTextRun,
        default_properties: Option<TextRunProperties>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        return Self {
            run_properties: TextRunProperties::from_raw(
                raw.clone().run_properties,
                default_properties.clone(),
                drawing_relationship.clone(),
                defined_names.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_font.clone(),
                font_ref_color.clone(),
            ),
            text: raw.text.unwrap_or(String::new()),
        };
    }
}
