#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::effect::effect::XlsxEffect;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effect?view=openxml-3.0.1
///
/// This element specifies a reference to an existing effect container
///
/// * Container: refer to an effect container with the name specified
/// * Fill: refers to the fill effect
/// * Line: refers to the line effect
/// * FillLine: refers to the combined fill and line effects
/// * Children: refers to the combined effects from logical child shapes or text
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum EffectReferenceTypeValues {
    /// refer to an effect container with the name specified
    Container(String),

    /// fill: refers to the fill effect
    Fill,

    /// line: refers to the line effect
    Line,

    /// fillLine: refers to the combined fill and line effects
    FillLine,

    /// children: refers to the combined effects from logical child shapes or text
    Children,
}

impl EffectReferenceTypeValues {
    pub(crate) fn from_raw(raw: Option<XlsxEffect>) -> Option<Self> {
        let Some(raw) = raw else { return None };
        let Some(s) = raw.r#ref else { return None };
        return Some(match s.as_ref() {
            "fill" => Self::Fill,
            "line" => Self::Line,
            "fillLine" => Self::FillLine,
            "children" => Self::Children,
            s => Self::Container(s.to_string()),
        });
    }
}
