#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
///
/// * Dash
/// * DashDot
/// * Dot
/// * LargeDash
/// * LargeDashDot
/// * LargeDashDotDot
/// * Solid
/// * SystemDash
/// * SystemDashDot
/// * SystemDashDotDot
/// * SystemDot
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetLineDashValues {
    Dash,
    DashDot,
    Dot,
    LargeDash,
    LargeDashDot,
    LargeDashDotDot,
    Solid,
    SystemDash,
    SystemDashDot,
    SystemDashDotDot,
    SystemDot,
}

impl PresetLineDashValues {
    pub(crate) fn default() -> Self {
        Self::Solid
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "dash" => Self::Dash,
            "dashDot" => Self::DashDot,
            "dot" => Self::Dot,
            "lgDash" => Self::LargeDash,
            "lgDashDot" => Self::LargeDashDot,
            "lgDashDotDot" => Self::LargeDashDotDot,
            "solid" => Self::Solid,
            "sysDash" => Self::SystemDash,
            "sysDashDot" => Self::SystemDashDot,
            "sysDashDotDot" => Self::SystemDashDotDot,
            "sysDot" => Self::SystemDot,
            _ => Self::default(),
        };
    }
}
