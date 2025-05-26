#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    raw::drawing::{
        effect::inner_shadow::XlsxInnerShadow,
        scheme::color_scheme::XlsxColorScheme,
        st_types::{emu_to_pt, st_angle_to_degree},
    },
};

/// innerShdw (Inner Shadow Effect)
///
/// specifies an inner shadow effect.
/// A shadow is applied within the edges of the object according to the parameters given by the attributes
///
///  Example:
/// ```
/// <a:innerShdw blurRad="10" dir"90" dist="10">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:innerShdw>
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.innershadow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct InnerShadow {
    /// color
    pub color: HexColor,

    /// blur radius in points
    pub blur_radius: f64,

    /// Specifies the direction to offset the shadow as angle
    pub direction: f64,

    /// Specifies how far to offset the shadow in points
    pub distance: f64,
}

impl InnerShadow {
    pub(crate) fn from_raw(
        raw: Option<XlsxInnerShadow>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(color) = raw.clone().color else {
            return None;
        };
        let Some(hex) = color.to_hex(color_scheme.clone(), ref_color.clone()) else {
            return None;
        };

        return Some(Self {
            color: hex,
            blur_radius: emu_to_pt(raw.clone().blur_rad.unwrap_or(0) as i64),
            direction: st_angle_to_degree(raw.clone().dir.unwrap_or(0) as i64),
            distance: emu_to_pt(raw.clone().dist.unwrap_or(0) as i64),
        });
    }
}
