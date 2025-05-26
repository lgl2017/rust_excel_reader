#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{shape::bevel::XlsxBevel, st_types::emu_to_pt};

/// - bevel: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevel?view=openxml-3.0.1
/// - bevel bottom: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelbottom?view=openxml-3.0.1
/// - bevel top: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.beveltop?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:bevelT w="254000" h="254000"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Bevel {
    /// Specifies the preset bevel type which defines the look of the bevel.
    pub preset: BevelPresetValues,

    /// Specifies the height of the bevel, or how far above the shape it is applied.
    pub height: f64,

    /// Specifies the width of the bevel, or how far into the shape it is applied.
    pub width: f64,
}

impl Bevel {
    pub(crate) fn from_raw(raw: Option<XlsxBevel>) -> Option<Self> {
        let Some(raw) = raw else { return None };
        if raw.h.is_none() && raw.w.is_none() && raw.prst.is_none() {
            return None;
        }

        let preset = BevelPresetValues::from_string(raw.clone().prst);
        let (mut width, mut height) = Self::w_h_from_preset(preset.clone());

        if let Some(w) = raw.clone().w {
            width = emu_to_pt(w as i64);
        };

        if let Some(h) = raw.clone().h {
            height = emu_to_pt(h as i64);
        };

        return Some(Self {
            preset,
            height,
            width,
        });
    }

    fn w_h_from_preset(preset: BevelPresetValues) -> (f64, f64) {
        match preset {
            BevelPresetValues::Angle => (6.0, 6.0),
            BevelPresetValues::ArtDeco => (9.0, 6.0),
            BevelPresetValues::Circle => (6.0, 6.0),
            BevelPresetValues::Convex => (6.0, 6.0),
            BevelPresetValues::CoolSlant => (13.0, 6.0),
            BevelPresetValues::Cross => (11.0, 6.0),
            BevelPresetValues::Divot => (11.0, 11.0),
            BevelPresetValues::HardEdge => (9.0, 6.0),
            BevelPresetValues::RelaxedInset => (6.0, 6.0),
            BevelPresetValues::Riblet => (8.0, 6.0),
            BevelPresetValues::Slope => (6.0, 6.0),
            BevelPresetValues::SoftRound => (12.0, 4.0),
        }
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelpresetvalues?view=openxml-3.0.1
///
/// * Angle
/// * ArtDeco
/// * Circle
/// * Convex
/// * CoolSlant
/// * Cross
/// * Divot
/// * HardEdge
/// * RelaxedInset
/// * Riblet
/// * Slope
/// * SoftRound
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BevelPresetValues {
    Angle,
    ArtDeco,
    Circle,
    Convex,
    CoolSlant,
    Cross,
    Divot,
    HardEdge,
    RelaxedInset,
    Riblet,
    Slope,
    SoftRound,
}

impl BevelPresetValues {
    pub(crate) fn default() -> Self {
        Self::Circle
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "angle" => Self::Angle,
            "artDeco" => Self::ArtDeco,
            "circle" => Self::Circle,
            "convex" => Self::Convex,
            "coolSlant" => Self::CoolSlant,
            "cross" => Self::Cross,
            "divot" => Self::Divot,
            "hardEdge" => Self::HardEdge,
            "relaxedInset" => Self::RelaxedInset,
            "riblet" => Self::Riblet,
            "slope" => Self::Slope,
            "softRound" => Self::SoftRound,
            _ => Self::default(),
        };
    }
}
