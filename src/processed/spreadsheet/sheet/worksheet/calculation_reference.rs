#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.referencemodevalues?view=openxml-3.0.1
///
/// * A1
/// * R1C1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum CalculationReferenceMode {
    A1,
    R1C1,
}

impl CalculationReferenceMode {
    pub(crate) fn default() -> Self {
        Self::A1
    }

    pub(crate) fn from_string(s: Option<String>) -> Option<Self> {
        let Some(s) = s else { return None };
        return match s.as_ref() {
            "A1" => Some(Self::A1),
            "R1C1" => Some(Self::R1C1),
            _ => None,
        };
    }
}
