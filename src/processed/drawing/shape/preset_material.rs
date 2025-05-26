#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelpresetvalues?view=openxml-3.0.1
///
/// * Clear
/// * DarkEdge
/// * Flat
/// * LegacyMatte
/// * LegacyMetal
/// * LegacyPlastic
/// * LegacyWireframe
/// * Matte
/// * Metal
/// * Plastic
/// * Powder
/// * SoftEdge
/// * SoftMetal
/// * TranslucentPowder
/// * WarmMatte
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetMaterialTypeValues {
    Clear,
    DarkEdge,
    Flat,
    LegacyMatte,
    LegacyMetal,
    LegacyPlastic,
    LegacyWireframe,
    Matte,
    Metal,
    Plastic,
    Powder,
    SoftEdge,
    SoftMetal,
    TranslucentPowder,
    WarmMatte,
}

impl PresetMaterialTypeValues {
    pub(crate) fn default() -> Self {
        Self::WarmMatte
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "clear" => Self::Clear,
            "dkEdge" => Self::DarkEdge,
            "flat" => Self::Flat,
            "legacyMatte" => Self::LegacyMatte,
            "legacyMetal" => Self::LegacyMetal,
            "legacyPlastic" => Self::LegacyPlastic,
            "legacyWireframe" => Self::LegacyWireframe,
            "matte" => Self::Matte,
            "metal" => Self::Metal,
            "plastic" => Self::Plastic,
            "powder" => Self::Powder,
            "softEdge" => Self::SoftEdge,
            "softmetal" => Self::SoftMetal,
            "translucentPowder" => Self::TranslucentPowder,
            "warmMatte" => Self::WarmMatte,
            _ => Self::default(),
        };
    }
}
