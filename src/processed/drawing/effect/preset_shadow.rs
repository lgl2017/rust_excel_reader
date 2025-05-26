#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;
use crate::raw::drawing::st_types::{emu_to_pt, st_angle_to_degree};
use crate::{common_types::HexColor, raw::drawing::effect::preset_shadow::XlsxPresetShadow};

/// prstShdw (Preset Shadow)
///
/// specifies that a preset shadow is to be used.
///
/// Each preset shadow is equivalent to a specific outer shadow effect.
/// For each preset shadow, the color element, direction attribute, and distance attribute represent the color, direction, and distance parameters of the corresponding outer shadow.
/// Additionally, the rotateWithShape attribute of corresponding outer shadow is always false. Other non-default parameters of the outer shadow are dependent on the prst attribute
///
///  Example:
/// ```
/// <a:prstShdw dir"90" dist="10" prst="shdw19">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:prstShdw>
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetshadow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct PresetShadow {
    /// color
    pub color: HexColor,

    /// Specifies the direction to offset the shadow as angle
    pub direction: f64,

    /// Specifies how far to offset the shadow in points
    pub distance: f64,

    ///	Specifies which preset shadow to use.
    pub preset: PresetShadowValues,
}

impl PresetShadow {
    pub(crate) fn from_raw(
        raw: Option<XlsxPresetShadow>,
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

        let Some(preset) = PresetShadowValues::from_string(raw.clone().prst) else {
            return None;
        };

        return Some(Self {
            color: hex,
            direction: st_angle_to_degree(raw.clone().dir.unwrap_or(0)),
            distance: emu_to_pt(raw.clone().dist.unwrap_or(0) as i64),
            preset,
        });
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetshadowvalues?view=openxml-3.0.1
///
/// * BackCenterPerspectiveShadow
/// * BackLeftLongPerspectiveShadow
/// * BackLeftPerspectiveShadow
/// * BackRightLongPerspectiveShadow
/// * BackRightPerspectiveShadow
/// * BottomLeftDropShadow
/// * ottomRightDropShadow
/// * BottomRightSmallDropShadow
/// * FrontBottomShadow
/// * FrontLeftLongPerspectiveShadow
/// * FrontLeftPerspectiveShadow
/// * FrontRightLongPerspectiveShadow
/// * FrontRightPerspectiveShadow
/// * ThreeDimensionalInnerBoxShadow
/// * ThreeDimensionalOuterBoxShadow
/// * TopLeftDoubleDropShadow
/// * TopLeftDropShadow
/// * TopLeftLargeDropShadow
/// * TopLeftSmallDropShadow
/// * TopRightDropShadow
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetShadowValues {
    BackCenterPerspectiveShadow,
    BackLeftLongPerspectiveShadow,
    BackLeftPerspectiveShadow,
    BackRightLongPerspectiveShadow,
    BackRightPerspectiveShadow,
    BottomLeftDropShadow,
    BottomRightDropShadow,
    BottomRightSmallDropShadow,
    FrontBottomShadow,
    FrontLeftLongPerspectiveShadow,
    FrontLeftPerspectiveShadow,
    FrontRightLongPerspectiveShadow,
    FrontRightPerspectiveShadow,
    ThreeDimensionalInnerBoxShadow,
    ThreeDimensionalOuterBoxShadow,
    TopLeftDoubleDropShadow,
    TopLeftDropShadow,
    TopLeftLargeDropShadow,
    TopLeftSmallDropShadow,
    TopRightDropShadow,
}

impl PresetShadowValues {
    pub(crate) fn from_string(s: Option<String>) -> Option<Self> {
        let Some(s) = s else { return None };
        return match s.as_ref() {
            "shdw19" => Some(Self::BackCenterPerspectiveShadow),
            "shdw11" => Some(Self::BackLeftLongPerspectiveShadow),
            "shdw3" => Some(Self::BackLeftPerspectiveShadow),
            "shdw12" => Some(Self::BackRightLongPerspectiveShadow),
            "shdw4" => Some(Self::BackRightPerspectiveShadow),
            "shdw5" => Some(Self::BottomLeftDropShadow),
            "shdw6" => Some(Self::BottomRightDropShadow),
            "shdw14" => Some(Self::BottomRightSmallDropShadow),
            "shdw20" => Some(Self::FrontBottomShadow),
            "shdw15" => Some(Self::FrontLeftLongPerspectiveShadow),
            "shdw7" => Some(Self::FrontLeftPerspectiveShadow),
            "shdw16" => Some(Self::FrontRightLongPerspectiveShadow),
            "shdw8" => Some(Self::FrontRightPerspectiveShadow),
            "shdw18" => Some(Self::ThreeDimensionalInnerBoxShadow),
            "shdw17" => Some(Self::ThreeDimensionalOuterBoxShadow),
            "shdw13" => Some(Self::TopLeftDoubleDropShadow),
            "shdw1" => Some(Self::TopLeftDropShadow),
            "shdw10" => Some(Self::TopLeftLargeDropShadow),
            "shdw9" => Some(Self::TopLeftSmallDropShadow),
            "shdw2" => Some(Self::TopRightDropShadow),
            _ => None,
        };
    }
}
