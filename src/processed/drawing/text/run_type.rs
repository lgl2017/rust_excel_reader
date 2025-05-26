#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::raw::drawing::text::paragraph::text_paragraphs::XlsxRunType;

use super::{line_break::LineBreak, text_field::TextField, text_run::TextRun};
use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::scheme::color_scheme::XlsxColorScheme,
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{font::Font, text_run_properties::TextRunProperties};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum RunTypeValues {
    /// TextRun
    Text(TextRun),

    /// TextLineBreak
    LineBreak(LineBreak),

    /// TextField
    ///
    /// contains generated text that the application should update periodically.
    TextField(TextField),
}

impl RunTypeValues {
    pub(crate) fn from_raw(
        raw: XlsxRunType,
        default_properties: Option<TextRunProperties>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        match raw {
            XlsxRunType::Text(r) => Self::Text(TextRun::from_raw(
                r,
                default_properties,
                drawing_relationship,
                defined_names,
                image_bytes,
                color_scheme,
                ref_font,
                font_ref_color,
            )),
            XlsxRunType::LineBreak(r) => Self::LineBreak(LineBreak::from_raw(
                r,
                default_properties,
                drawing_relationship,
                defined_names,
                image_bytes,
                color_scheme,
                ref_font,
                font_ref_color,
            )),
            XlsxRunType::TextField(r) => Self::TextField(TextField::from_raw(
                r,
                default_properties,
                drawing_relationship,
                defined_names,
                image_bytes,
                color_scheme,
                ref_font,
                font_ref_color,
            )),
        }
    }
}
