#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.compoundlinevalues?view=openxml-3.0.1
///
/// * Double
/// * Single
/// * ThickThin
/// * ThinThick
/// * Triple
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum CompoundLineValues {
    Double,
    Single,
    ThickThin,
    ThinThick,
    Triple,
}

impl CompoundLineValues {
    pub(crate) fn default() -> Self {
        Self::Single
    }

    pub(crate) fn from_string(s: Option<String>, reference: Option<Self>) -> Self {
        let Some(s) = s else {
            return reference.unwrap_or(Self::default());
        };
        return match s.as_ref() {
            "dbl" => Self::Double,
            "sng" => Self::Single,
            "thickThin" => Self::ThickThin,
            "thinThick" => Self::ThinThick,
            "tri" => Self::Triple,
            _ => Self::default(),
        };
    }
}
