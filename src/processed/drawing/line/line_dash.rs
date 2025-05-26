#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::line::outline::XlsxOutline;

use super::{dash_stop::DashStop, preset_line_dash::PresetLineDashValues};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LineDashTypeValues {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customdash?view=openxml-3.0.1
    ///
    /// This element specifies a custom dashing scheme.
    /// It is a list of dash stop elements which represent building block atoms upon which the custom dashing scheme is built.
    CustomDash(Vec<DashStop>),

    /// Preset dash: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
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
    PresetDash(PresetLineDashValues),
}

impl LineDashTypeValues {
    pub(crate) fn default() -> Self {
        Self::PresetDash(PresetLineDashValues::default())
    }

    pub(crate) fn from_raw(raw: XlsxOutline, reference: Option<Self>) -> Self {
        if let Some(preset) = raw.clone().present_dash {
            return Self::PresetDash(PresetLineDashValues::from_string(Some(preset)));
        };

        if let Some(custom_dash) = raw.clone().custom_dash {
            let dash_stop = custom_dash
                .ds
                .unwrap_or(vec![])
                .into_iter()
                .map(|d| DashStop::from_raw(d))
                .collect();

            return Self::CustomDash(dash_stop);
        };

        return reference.unwrap_or(Self::default());
    }
}
