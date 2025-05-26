#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{scene::rotation::XlsxRotation, st_types::st_angle_to_degree};

use super::camera::PresetCameraValues;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rotation?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:rot lat="0" lon="0" rev="6000000"/>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Rotation {
    /// long: longitude coordinate, (x rotation)
    pub x_rotation: f64,

    /// lat: latitude coordinate (y rotation)
    pub y_rotation: f64,

    /// rev: revolution about the axis as the latitude and longitude coordinates
    ///
    /// - z rotation in terms of 3d rotation
    /// - angle in terms of lighting
    pub z_rotation: f64,
}

impl Rotation {
    pub(crate) fn from_light_rig(raw: Option<XlsxRotation>) -> Self {
        let Some(raw) = raw else {
            return Self {
                x_rotation: 0.0,
                y_rotation: 0.0,
                z_rotation: 0.0,
            };
        };
        return Self {
            x_rotation: st_angle_to_degree(raw.clone().long.unwrap_or(0)),
            y_rotation: st_angle_to_degree(raw.clone().lat.unwrap_or(0)),
            z_rotation: st_angle_to_degree(raw.clone().rev.unwrap_or(0)),
        };
    }
    pub(crate) fn from_camera(
        raw: Option<XlsxRotation>,
        preset: Option<PresetCameraValues>,
    ) -> Option<Self> {
        if let Some(rot) = raw {
            return Some(Self {
                x_rotation: st_angle_to_degree(rot.long.unwrap_or(0)),
                y_rotation: st_angle_to_degree(rot.lat.unwrap_or(0)),
                z_rotation: st_angle_to_degree(rot.rev.unwrap_or(0)),
            });
        }
        let Some(preset) = preset else {
            return None;
        };
        match preset {
            PresetCameraValues::IsometricBottomDown => Self::build_self_helper(314.7, 35.4, 299.8),
            PresetCameraValues::IsometricBottomUp => None,
            PresetCameraValues::IsometricLeftDown => Self::build_self_helper(45.0, 35.0, 0.0),
            PresetCameraValues::IsometricLeftUp => None,
            PresetCameraValues::IsometricOffAxis1Left => Self::build_self_helper(64.0, 18.0, 0.0),
            PresetCameraValues::IsometricOffAxis1Right => Self::build_self_helper(334.0, 18.0, 0.0),
            PresetCameraValues::IsometricOffAxis1Top => Self::build_self_helper(306.5, 301.3, 57.6),
            PresetCameraValues::IsometricOffAxis2Left => Self::build_self_helper(26.0, 18.0, 0.0),
            PresetCameraValues::IsometricOffAxis2Right => Self::build_self_helper(296.0, 18.0, 0.0),
            PresetCameraValues::IsometricOffAxis2Top => Self::build_self_helper(53.5, 301.3, 302.4),
            PresetCameraValues::IsometricOffAxis3Bottom => None,
            PresetCameraValues::IsometricOffAxis3Left => None,
            PresetCameraValues::IsometricOffAxis3Right => None,
            PresetCameraValues::IsometricOffAxis4Bottom => None,
            PresetCameraValues::IsometricOffAxis4Left => None,
            PresetCameraValues::IsometricOffAxis4Right => None,
            PresetCameraValues::IsometricRightDown => None,
            PresetCameraValues::IsometricRightUp => Self::build_self_helper(315.0, 35.0, 0.0),
            PresetCameraValues::IsometricTopDown => None,
            PresetCameraValues::IsometricTopUp => Self::build_self_helper(314.7, 324.6, 60.2),
            PresetCameraValues::LegacyObliqueBottom => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueBottomLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueBottomRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueFront => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueTop => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueTopLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyObliqueTopRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyPerspectiveBottom => Self::build_self_helper(0.0, 20.0, 0.0),
            PresetCameraValues::LegacyPerspectiveBottomLeft => None,
            PresetCameraValues::LegacyPerspectiveBottomRight => None,
            PresetCameraValues::LegacyPerspectiveFront => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::LegacyPerspectiveLeft => Self::build_self_helper(20.0, 0.0, 0.0),
            PresetCameraValues::LegacyPerspectiveRight => Self::build_self_helper(340.0, 0.0, 0.0),
            PresetCameraValues::LegacyPerspectiveTop => Self::build_self_helper(0.0, 340.0, 0.0),
            PresetCameraValues::LegacyPerspectiveTopLeft => None,
            PresetCameraValues::LegacyPerspectiveTopRight => None,
            PresetCameraValues::ObliqueBottom => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueBottomLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueBottomRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueTop => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueTopLeft => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::ObliqueTopRight => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::OrthographicFront => None,
            PresetCameraValues::PerspectiveAbove => Self::build_self_helper(0.0, 340.0, 0.0),
            PresetCameraValues::PerspectiveAboveLeftFacing => None,
            PresetCameraValues::PerspectiveAboveRightFacing => None,
            PresetCameraValues::PerspectiveBelow => Self::build_self_helper(0.0, 20.0, 0.0),
            PresetCameraValues::PerspectiveContrastingLeftFacing => {
                Self::build_self_helper(43.9, 10.4, 356.4)
            }
            PresetCameraValues::PerspectiveContrastingRightFacing => {
                Self::build_self_helper(316.1, 10.4, 3.6)
            }
            PresetCameraValues::PerspectiveFront => Self::build_self_helper(0.0, 0.0, 0.0),
            PresetCameraValues::PerspectiveHeroicExtremeLeftFacing => {
                Self::build_self_helper(34.5, 8.1, 357.1)
            }
            PresetCameraValues::PerspectiveHeroicExtremeRightFacing => {
                Self::build_self_helper(325.5, 8.1, 2.9)
            }
            PresetCameraValues::PerspectiveHeroicLeftFacing => None,
            PresetCameraValues::PerspectiveHeroicRightFacing => None,
            PresetCameraValues::PerspectiveLeft => Self::build_self_helper(20.0, 0.0, 0.0),
            PresetCameraValues::PerspectiveRelaxed => Self::build_self_helper(0.0, 309.6, 0.0),
            PresetCameraValues::PerspectiveRelaxedModerately => {
                Self::build_self_helper(0.0, 324.8, 0.0)
            }
            PresetCameraValues::PerspectiveRight => Self::build_self_helper(340.0, 0.0, 0.0),
        }
    }

    fn build_self_helper(x: f64, y: f64, z: f64) -> Option<Self> {
        return Some(Self {
            x_rotation: x,
            y_rotation: y,
            z_rotation: z,
        });
    }
}
