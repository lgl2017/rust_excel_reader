#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.penalignmentvalues?view=openxml-3.0.1
///
/// * Center
/// * Insert
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PenAlignmentValues {
    Center,
    Insert,
}

impl PenAlignmentValues {
    pub(crate) fn from_string(s: Option<String>, reference: Option<Self>) -> Self {
        let Some(s) = s else {
            return reference.unwrap_or(Self::Center);
        };
        return match s.as_ref() {
            "ctr" => Self::Center,
            "in" => Self::Insert,
            _ => Self::Center,
        };
    }
}
