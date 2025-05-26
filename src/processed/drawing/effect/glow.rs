#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    raw::drawing::{
        effect::glow::XlsxGlow, scheme::color_scheme::XlsxColorScheme, st_types::emu_to_pt,
    },
};

/// glow (Glow Effect)
///
/// specifies a glow effect, in which a color blurred outline is added outside the edges of the object.
///
/// Example:
/// ```
/// <a:glow rad="10">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:glow>
/// ```
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.glow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Glow {
    /// color of the glow
    pub color: HexColor,

    /// radius of the glow in points
    pub radius: f64,
}

impl Glow {
    pub(crate) fn from_raw(
        raw: Option<XlsxGlow>,
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
            radius: emu_to_pt(raw.clone().rad.unwrap_or(0) as i64),
        });
    }
}
