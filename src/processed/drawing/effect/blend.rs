#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::drawing::{effect::blend::XlsxBlend, scheme::color_scheme::XlsxColorScheme},
};

use super::effect_container::EffectContainer;

/// blend (Blend Effect)
///
/// Specifies a blend of several effects.
/// The container specifies the raw effects to blend while the blend mode specifies how the effects are to be blended.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blend?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Blend {
    /// specifies the raw effects to blend
    pub effect: Box<EffectContainer>,

    /// Specifies how to blend the two effects
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blendmodevalues?view=openxml-3.0.1
    pub blend_mode: BlendModeValues,
}

impl Blend {
    pub(crate) fn from_raw(
        raw: Option<XlsxBlend>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else {
            return None;
        };

        let Some(container) = EffectContainer::from_raw(
            raw.cont,
            drawing_relationship,
            image_bytes,
            color_scheme.clone(),
            ref_color.clone(),
        ) else {
            return None;
        };
        return Some(Self {
            effect: container,
            blend_mode: BlendModeValues::from_raw(raw.blend),
        });
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blendmodevalues?view=openxml-3.0.1
///
/// * Darken
/// * Lighten
/// * Multiply
/// * Overlay
/// * Screen
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BlendModeValues {
    Darken,
    Lighten,
    Multiply,
    Overlay,
    Screen,
}

impl BlendModeValues {
    pub(crate) fn from_raw(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Overlay };
        return match s.as_ref() {
            "darken" => Self::Darken,
            "lighten" => Self::Lighten,
            "mult" => Self::Multiply,
            "over" => Self::Overlay,
            "screen" => Self::Screen,
            _ => Self::Overlay,
        };
    }
}
