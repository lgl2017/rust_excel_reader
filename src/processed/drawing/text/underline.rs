use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use super::text_underline_values::TextUnderlineValues;
use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::{fill::Fill, line::outline::Outline},
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme,
        text::default_text_run_properties::XlsxTextRunProperties,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
///
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Underline {
    /// underline type
    pub r#type: TextUnderlineValues,

    /// fill of underline
    ///
    /// (when not specified, underline for a run of text should be of the same color as the text run within which it is contained.)
    pub fill: Fill,

    /// Secifies the properties for the stroke of the underline that is present within a run of text.
    ///
    /// (when not specified, the stroke style of an underline for a run of text should be of the same as the text run within which it is contained.)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub outline: Option<Outline>,
}

impl Underline {
    pub(crate) fn default() -> Self {
        Self {
            fill: Fill::SolidFill("000000ff".to_string()),
            outline: None,
            r#type: TextUnderlineValues::default(),
        }
    }

    pub(crate) fn from_text_run_properties(
        raw: XlsxTextRunProperties,
        default_underline: Option<Self>,
        text_outline: Option<Outline>,
        text_fill: Fill,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        let fill = if let Some(f) = Fill::from_raw(
            raw.clone().underline_fill,
            None,
            drawing_relationship.clone(),
            image_bytes.clone(),
            color_scheme.clone(),
            font_ref_color.clone(),
        ) {
            f
        } else {
            if let Some(default_underline) = default_underline.clone() {
                default_underline.fill
            } else {
                text_fill
            }
        };

        let outline = if let Some(outline) =
            Outline::from_raw(raw.clone().underline_stroke, color_scheme.clone(), None)
        {
            Some(outline)
        } else {
            if let Some(default_underline) = default_underline.clone() {
                default_underline.outline
            } else {
                text_outline
            }
        };

        let underline = if let Some(u) = raw.clone().underline {
            TextUnderlineValues::from_string(Some(u))
        } else {
            if let Some(default_underline) = default_underline.clone() {
                default_underline.r#type
            } else {
                TextUnderlineValues::default()
            }
        };

        return Self {
            fill,
            outline,
            r#type: underline,
        };
    }
}
