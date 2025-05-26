#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::scene::light_rig::XlsxLightRig;

use super::rotation::Rotation;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrig?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:lightRig rig="twoPt" dir="t">
///     <a:rot lat="0" lon="0" rev="6000000"/>
/// </a:lightRig>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct LightRig {
    /// Rig Preset
    pub preset: LightRigPresetValues,

    /// Direction
    pub direction: LightRigDirectionValues,

    /// rotation
    pub rotation: Rotation,
}

impl LightRig {
    pub(crate) fn from_raw(raw: Option<XlsxLightRig>) -> Option<Self> {
        let Some(raw) = raw else { return None };
        return Some(Self {
            preset: LightRigPresetValues::from_string(raw.clone().rig),
            direction: LightRigDirectionValues::from_string(raw.clone().dir),
            rotation: Rotation::from_light_rig(raw.clone().rot),
        });
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrigvalues?view=openxml-3.0.1
///
/// * Balanced
/// * BrightRoom
/// * Chilly
/// * Contrasting
/// * Flat
/// * Flood
/// * Freezing
/// * Glow
/// * Harsh
/// * LegacyFlat1
/// * LegacyFlat2
/// * LegacyFlat3
/// * LegacyFlat4
/// * LegacyHarsh1
/// * LegacyHarsh2
/// * LegacyHarsh3
/// * LegacyHarsh4
/// * LegacyNormal1
/// * LegacyNormal2
/// * LegacyNormal3
/// * LegacyNormal4
/// * Morning
/// * Soft
/// * Sunrise
/// * Sunset
/// * ThreePoints
/// * TwoPoints
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LightRigPresetValues {
    Balanced,
    BrightRoom,
    Chilly,
    Contrasting,
    Flat,
    Flood,
    Freezing,
    Glow,
    Harsh,
    LegacyFlat1,
    LegacyFlat2,
    LegacyFlat3,
    LegacyFlat4,
    LegacyHarsh1,
    LegacyHarsh2,
    LegacyHarsh3,
    LegacyHarsh4,
    LegacyNormal1,
    LegacyNormal2,
    LegacyNormal3,
    LegacyNormal4,
    Morning,
    Soft,
    Sunrise,
    Sunset,
    ThreePoints,
    TwoPoints,
}

impl LightRigPresetValues {
    pub(crate) fn default() -> Self {
        Self::ThreePoints
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "balanced" => Self::Balanced,
            "brightRoom" => Self::BrightRoom,
            "chilly" => Self::Chilly,
            "contrasting" => Self::Contrasting,
            "flat" => Self::Flat,
            "flood" => Self::Flood,
            "freezing" => Self::Freezing,
            "glow" => Self::Glow,
            "harsh" => Self::Harsh,
            "legacyFlat1" => Self::LegacyFlat1,
            "legacyFlat2" => Self::LegacyFlat2,
            "legacyFlat3" => Self::LegacyFlat3,
            "legacyFlat4" => Self::LegacyFlat4,
            "legacyHarsh1" => Self::LegacyHarsh1,
            "legacyHarsh2" => Self::LegacyHarsh2,
            "legacyHarsh3" => Self::LegacyHarsh3,
            "legacyHarsh4" => Self::LegacyHarsh4,
            "legacyNormal1" => Self::LegacyNormal1,
            "legacyNormal2" => Self::LegacyNormal2,
            "legacyNormal3" => Self::LegacyNormal3,
            "legacyNormal4" => Self::LegacyNormal4,
            "morning" => Self::Morning,
            "soft" => Self::Soft,
            "sunrise" => Self::Sunrise,
            "sunset" => Self::Sunset,
            "threePt" => Self::ThreePoints,
            "twoPt" => Self::TwoPoints,
            _ => Self::default(),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lightrigdirectionvalues?view=openxml-3.0.1
///
/// * Bottom
/// * BottomLeft
/// * BottomRight
/// * Left
/// * Right
/// * Top
/// * TopLeft
/// * TopRight
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LightRigDirectionValues {
    Bottom,
    BottomLeft,
    BottomRight,
    Left,
    Right,
    Top,
    TopLeft,
    TopRight,
}

impl LightRigDirectionValues {
    pub(crate) fn default() -> Self {
        Self::Top
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "b" => Self::Bottom,
            "bl" => Self::BottomLeft,
            "br" => Self::BottomRight,
            "l" => Self::Left,
            "r" => Self::Right,
            "t" => Self::Top,
            "tl" => Self::TopLeft,
            "tr" => Self::TopRight,
            _ => Self::default(),
        };
    }
}
