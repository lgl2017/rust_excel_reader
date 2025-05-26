#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{scene::camera::XlsxCamera, st_types::st_angle_to_degree};

use super::rotation::Rotation;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.camera?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:camera prst="orthographicFront">
///     <a:rot lat="19902513" lon="17826689" rev="1362739"/>
/// </a:camera>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Camera {
    /// Preset Camera Type
    pub preset: PresetCameraValues,

    /// rotation
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub rotation: Option<Rotation>,

    /// fov: Field of view (Perspective)
    pub perspective: f64,
    // /// Zoom
    // pub zoom: Option<i64>,
}

impl Camera {
    pub(crate) fn from_raw(raw: Option<XlsxCamera>) -> Option<Self> {
        let Some(raw) = raw else { return None };
        if raw.prst.is_none() && raw.rot.is_none() {
            return None;
        };

        let preset = PresetCameraValues::from_string(raw.clone().prst);
        let rotation = Rotation::from_camera(raw.clone().rot, Some(preset.clone()));

        let perspective = if let Some(p) = raw.clone().fov {
            st_angle_to_degree(p)
        } else {
            Self::perspective_from_preset(preset.clone())
        };
        return Some(Self {
            preset,
            rotation,
            perspective,
        });
    }

    fn perspective_from_preset(preset: PresetCameraValues) -> f64 {
        match preset {
            PresetCameraValues::IsometricBottomDown => 0.0,
            PresetCameraValues::IsometricBottomUp => 0.0,
            PresetCameraValues::IsometricLeftDown => 0.0,
            PresetCameraValues::IsometricLeftUp => 0.0,
            PresetCameraValues::IsometricOffAxis1Left => 0.0,
            PresetCameraValues::IsometricOffAxis1Right => 0.0,
            PresetCameraValues::IsometricOffAxis1Top => 0.0,
            PresetCameraValues::IsometricOffAxis2Left => 0.0,
            PresetCameraValues::IsometricOffAxis2Right => 0.0,
            PresetCameraValues::IsometricOffAxis2Top => 0.0,
            PresetCameraValues::IsometricOffAxis3Bottom => 0.0,
            PresetCameraValues::IsometricOffAxis3Left => 0.0,
            PresetCameraValues::IsometricOffAxis3Right => 0.0,
            PresetCameraValues::IsometricOffAxis4Bottom => 0.0,
            PresetCameraValues::IsometricOffAxis4Left => 0.0,
            PresetCameraValues::IsometricOffAxis4Right => 0.0,
            PresetCameraValues::IsometricRightDown => 0.0,
            PresetCameraValues::IsometricRightUp => 0.0,
            PresetCameraValues::IsometricTopDown => 0.0,
            PresetCameraValues::IsometricTopUp => 0.0,
            PresetCameraValues::LegacyObliqueBottom => 0.0,
            PresetCameraValues::LegacyObliqueBottomLeft => 0.0,
            PresetCameraValues::LegacyObliqueBottomRight => 0.0,
            PresetCameraValues::LegacyObliqueFront => 0.0,
            PresetCameraValues::LegacyObliqueLeft => 0.0,
            PresetCameraValues::LegacyObliqueRight => 0.0,
            PresetCameraValues::LegacyObliqueTop => 0.0,
            PresetCameraValues::LegacyObliqueTopLeft => 0.0,
            PresetCameraValues::LegacyObliqueTopRight => 0.0,
            PresetCameraValues::LegacyPerspectiveBottom => 45.0,
            PresetCameraValues::LegacyPerspectiveBottomLeft => 0.0,
            PresetCameraValues::LegacyPerspectiveBottomRight => 0.0,
            PresetCameraValues::LegacyPerspectiveFront => 45.0,
            PresetCameraValues::LegacyPerspectiveLeft => 45.0,
            PresetCameraValues::LegacyPerspectiveRight => 45.0,
            PresetCameraValues::LegacyPerspectiveTop => 45.0,
            PresetCameraValues::LegacyPerspectiveTopLeft => 0.0,
            PresetCameraValues::LegacyPerspectiveTopRight => 0.0,
            PresetCameraValues::ObliqueBottom => 0.0,
            PresetCameraValues::ObliqueBottomLeft => 0.0,
            PresetCameraValues::ObliqueBottomRight => 0.0,
            PresetCameraValues::ObliqueLeft => 0.0,
            PresetCameraValues::ObliqueRight => 0.0,
            PresetCameraValues::ObliqueTop => 0.0,
            PresetCameraValues::ObliqueTopLeft => 0.0,
            PresetCameraValues::ObliqueTopRight => 0.0,
            PresetCameraValues::OrthographicFront => 0.0,
            PresetCameraValues::PerspectiveAbove => 45.0,
            PresetCameraValues::PerspectiveAboveLeftFacing => 0.0,
            PresetCameraValues::PerspectiveAboveRightFacing => 0.0,
            PresetCameraValues::PerspectiveBelow => 45.0,
            PresetCameraValues::PerspectiveContrastingLeftFacing => 45.0,
            PresetCameraValues::PerspectiveContrastingRightFacing => 45.0,
            PresetCameraValues::PerspectiveFront => 45.0,
            PresetCameraValues::PerspectiveHeroicExtremeLeftFacing => 80.0,
            PresetCameraValues::PerspectiveHeroicExtremeRightFacing => 80.0,
            PresetCameraValues::PerspectiveHeroicLeftFacing => 0.0,
            PresetCameraValues::PerspectiveHeroicRightFacing => 0.0,
            PresetCameraValues::PerspectiveLeft => 45.0,
            PresetCameraValues::PerspectiveRelaxed => 45.0,
            PresetCameraValues::PerspectiveRelaxedModerately => 45.0,
            PresetCameraValues::PerspectiveRight => 45.0,
        }
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetcameravalues?view=openxml-3.0.1
///
/// * IsometricBottomDown
/// * IsometricBottomUp
/// * IsometricLeftDown
/// * IsometricLeftUp
/// * IsometricOffAxis1Left
/// * IsometricOffAxis1Right
/// * IsometricOffAxis1Top
/// * IsometricOffAxis2Left
/// * IsometricOffAxis2Right
/// * IsometricOffAxis2Top
/// * IsometricOffAxis3Bottom
/// * IsometricOffAxis3Left
/// * IsometricOffAxis3Right
/// * IsometricOffAxis4Bottom
/// * IsometricOffAxis4Left
/// * IsometricOffAxis4Right
/// * IsometricRightDown
/// * IsometricRightUp
/// * IsometricTopDown
/// * IsometricTopUp
/// * LegacyObliqueBottom
/// * LegacyObliqueBottomLeft
/// * LegacyObliqueBottomRight
/// * LegacyObliqueFront
/// * LegacyObliqueLeft
/// * LegacyObliqueRight
/// * LegacyObliqueTop
/// * LegacyObliqueTopLeft
/// * LegacyObliqueTopRight
/// * LegacyPerspectiveBottom
/// * LegacyPerspectiveBottomLeft
/// * LegacyPerspectiveBottomRight
/// * LegacyPerspectiveFront
/// * LegacyPerspectiveLeft
/// * LegacyPerspectiveRight
/// * LegacyPerspectiveTop
/// * LegacyPerspectiveTopLeft
/// * LegacyPerspectiveTopRight
/// * ObliqueBottom
/// * ObliqueBottomLeft
/// * ObliqueBottomRight
/// * ObliqueLeft
/// * ObliqueRight
/// * ObliqueTop
/// * ObliqueTopLeft
/// * ObliqueTopRight
/// * OrthographicFront
/// * PerspectiveAbove
/// * PerspectiveAboveLeftFacing
/// * PerspectiveAboveRightFacing
/// * PerspectiveBelow
/// * PerspectiveContrastingLeftFacing
/// * PerspectiveContrastingRightFacing
/// * PerspectiveFront
/// * PerspectiveHeroicExtremeLeftFacing
/// * PerspectiveHeroicExtremeRightFacing
/// * PerspectiveHeroicLeftFacing
/// * PerspectiveHeroicRightFacing
/// * PerspectiveLeft
/// * PerspectiveRelaxed
/// * PerspectiveRelaxedModerately
/// * PerspectiveRight
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetCameraValues {
    IsometricBottomDown,
    IsometricBottomUp,
    IsometricLeftDown,
    IsometricLeftUp,
    IsometricOffAxis1Left,
    IsometricOffAxis1Right,
    IsometricOffAxis1Top,
    IsometricOffAxis2Left,
    IsometricOffAxis2Right,
    IsometricOffAxis2Top,
    IsometricOffAxis3Bottom,
    IsometricOffAxis3Left,
    IsometricOffAxis3Right,
    IsometricOffAxis4Bottom,
    IsometricOffAxis4Left,
    IsometricOffAxis4Right,
    IsometricRightDown,
    IsometricRightUp,
    IsometricTopDown,
    IsometricTopUp,
    LegacyObliqueBottom,
    LegacyObliqueBottomLeft,
    LegacyObliqueBottomRight,
    LegacyObliqueFront,
    LegacyObliqueLeft,
    LegacyObliqueRight,
    LegacyObliqueTop,
    LegacyObliqueTopLeft,
    LegacyObliqueTopRight,
    LegacyPerspectiveBottom,
    LegacyPerspectiveBottomLeft,
    LegacyPerspectiveBottomRight,
    LegacyPerspectiveFront,
    LegacyPerspectiveLeft,
    LegacyPerspectiveRight,
    LegacyPerspectiveTop,
    LegacyPerspectiveTopLeft,
    LegacyPerspectiveTopRight,
    ObliqueBottom,
    ObliqueBottomLeft,
    ObliqueBottomRight,
    ObliqueLeft,
    ObliqueRight,
    ObliqueTop,
    ObliqueTopLeft,
    ObliqueTopRight,
    OrthographicFront,
    PerspectiveAbove,
    PerspectiveAboveLeftFacing,
    PerspectiveAboveRightFacing,
    PerspectiveBelow,
    PerspectiveContrastingLeftFacing,
    PerspectiveContrastingRightFacing,
    PerspectiveFront,
    PerspectiveHeroicExtremeLeftFacing,
    PerspectiveHeroicExtremeRightFacing,
    PerspectiveHeroicLeftFacing,
    PerspectiveHeroicRightFacing,
    PerspectiveLeft,
    PerspectiveRelaxed,
    PerspectiveRelaxedModerately,
    PerspectiveRight,
}

impl PresetCameraValues {
    pub(crate) fn default() -> Self {
        Self::OrthographicFront
    }

    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "isometricBottomDown" => Self::IsometricBottomDown,
            "isometricBottomUp" => Self::IsometricBottomUp,
            "isometricLeftDown" => Self::IsometricLeftDown,
            "isometricLeftUp" => Self::IsometricLeftUp,
            "isometricOffAxis1Left" => Self::IsometricOffAxis1Left,
            "isometricOffAxis1Right" => Self::IsometricOffAxis1Right,
            "isometricOffAxis1Top" => Self::IsometricOffAxis1Top,
            "isometricOffAxis2Left" => Self::IsometricOffAxis2Left,
            "isometricOffAxis2Right" => Self::IsometricOffAxis2Right,
            "isometricOffAxis2Top" => Self::IsometricOffAxis2Top,
            "isometricOffAxis3Bottom" => Self::IsometricOffAxis3Bottom,
            "isometricOffAxis3Left" => Self::IsometricOffAxis3Left,
            "isometricOffAxis3Right" => Self::IsometricOffAxis3Right,
            "isometricOffAxis4Bottom" => Self::IsometricOffAxis4Bottom,
            "isometricOffAxis4Left" => Self::IsometricOffAxis4Left,
            "isometricOffAxis4Right" => Self::IsometricOffAxis4Right,
            "isometricRightDown" => Self::IsometricRightDown,
            "isometricRightUp" => Self::IsometricRightUp,
            "isometricTopDown" => Self::IsometricTopDown,
            "isometricTopUp" => Self::IsometricTopUp,
            "legacyObliqueBottom" => Self::LegacyObliqueBottom,
            "legacyObliqueBottomLeft" => Self::LegacyObliqueBottomLeft,
            "legacyObliqueBottomRight" => Self::LegacyObliqueBottomRight,
            "legacyObliqueFront" => Self::LegacyObliqueFront,
            "legacyObliqueLeft" => Self::LegacyObliqueLeft,
            "legacyObliqueRight" => Self::LegacyObliqueRight,
            "legacyObliqueTop" => Self::LegacyObliqueTop,
            "legacyObliqueTopLeft" => Self::LegacyObliqueTopLeft,
            "legacyObliqueTopRight" => Self::LegacyObliqueTopRight,
            "legacyPerspectiveBottom" => Self::LegacyPerspectiveBottom,
            "legacyPerspectiveBottomLeft" => Self::LegacyPerspectiveBottomLeft,
            "legacyPerspectiveBottomRight" => Self::LegacyPerspectiveBottomRight,
            "legacyPerspectiveFront" => Self::LegacyPerspectiveFront,
            "legacyPerspectiveLeft" => Self::LegacyPerspectiveLeft,
            "legacyPerspectiveRight" => Self::LegacyPerspectiveRight,
            "legacyPerspectiveTop" => Self::LegacyPerspectiveTop,
            "legacyPerspectiveTopLeft" => Self::LegacyPerspectiveTopLeft,
            "legacyPerspectiveTopRight" => Self::LegacyPerspectiveTopRight,
            "obliqueBottom" => Self::ObliqueBottom,
            "obliqueBottomLeft" => Self::ObliqueBottomLeft,
            "obliqueBottomRight" => Self::ObliqueBottomRight,
            "obliqueLeft" => Self::ObliqueLeft,
            "obliqueRight" => Self::ObliqueRight,
            "obliqueTop" => Self::ObliqueTop,
            "obliqueTopLeft" => Self::ObliqueTopLeft,
            "obliqueTopRight" => Self::ObliqueTopRight,
            "orthographicFront" => Self::OrthographicFront,
            "perspectiveAbove" => Self::PerspectiveAbove,
            "perspectiveAboveLeftFacing" => Self::PerspectiveAboveLeftFacing,
            "perspectiveAboveRightFacing" => Self::PerspectiveAboveRightFacing,
            "perspectiveBelow" => Self::PerspectiveBelow,
            "perspectiveContrastingLeftFacing" => Self::PerspectiveContrastingLeftFacing,
            "perspectiveContrastingRightFacing" => Self::PerspectiveContrastingRightFacing,
            "perspectiveFront" => Self::PerspectiveFront,
            "perspectiveHeroicExtremeLeftFacing" => Self::PerspectiveHeroicExtremeLeftFacing,
            "perspectiveHeroicExtremeRightFacing" => Self::PerspectiveHeroicExtremeRightFacing,
            "perspectiveHeroicLeftFacing" => Self::PerspectiveHeroicLeftFacing,
            "perspectiveHeroicRightFacing" => Self::PerspectiveHeroicRightFacing,
            "perspectiveLeft" => Self::PerspectiveLeft,
            "perspectiveRelaxed" => Self::PerspectiveRelaxed,
            "perspectiveRelaxedModerately" => Self::PerspectiveRelaxedModerately,
            "perspectiveRight" => Self::PerspectiveRight,
            _ => Self::default(),
        };
    }
}
