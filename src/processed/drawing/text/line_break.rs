use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme, text::paragraph::line_break::XlsxTextLineBreak,
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{font::Font, text_run_properties::TextRunProperties};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.break?view=openxml-3.0.1
///
/// This element specifies the existence of a vertical line break between two runs of text within a paragraph.
/// In addition to specifying a vertical space between two runs of text, this element can also have run properties specified via the rPr child element.
/// This sets the formatting of text for the line break so that if text is later inserted there that a new run can be generated with the correct formatting.
///
/// Example:
/// ```
/// <a:br>
/// <a:rPr lang="en-US" sz="1100">
///     <a:ln>
///         <a:solidFill>
///             <a:schemeClr val="accent5">
///                 <a:alpha val="99339" />
///             </a:schemeClr>
///         </a:solidFill>
///     </a:ln>
/// </a:rPr>
/// </a:br>
/// ```
/// br (Text Line Break)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct LineBreak {
    // rPr (Text Run Properties)
    pub run_properties: TextRunProperties,
}

impl LineBreak {
    pub(crate) fn from_raw(
        raw: XlsxTextLineBreak,
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
                raw.clone().text_run_properties,
                default_properties.clone(),
                drawing_relationship.clone(),
                defined_names.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_font.clone(),
                font_ref_color.clone(),
            ),
        };
    }
}
