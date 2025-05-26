#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    scene::backdrop::{XlsxAnchor, XlsxBackDrop, XlsxVector},
    st_types::emu_to_pt,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backdrop?view=openxml-3.0.1
///
/// This element defines a plane in which effects, such as glow and shadow, are applied in relation to the shape they are being applied to.
/// The points and vectors contained within the backdrop define a plane in 3D space.
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct BackDrop {
    /// anchor:
    ///
    /// This element specifies a point in 3D space.
    /// This point is the point in space that anchors the backdrop plane.
    pub anchor: Anchor,

    /// normal vector:
    ///
    /// This element defines a normal vector.
    /// To be more precise, this attribute defines a vector normal to the face of the backdrop plane.
    pub normal: Vector,

    /// up vector:
    ///
    /// This element defines a vector representing up.
    /// To be more precise, this attribute defines a vector representing up in relation to the face of the backdrop plane.
    pub up: Vector,
}

impl BackDrop {
    pub(crate) fn from_raw(raw: Option<XlsxBackDrop>) -> Option<Self> {
        let Some(raw) = raw else { return None };
        return Some(Self {
            anchor: Anchor::from_raw(raw.clone().anchor),
            normal: Vector::from_raw(raw.clone().norm),
            up: Vector::from_raw(raw.clone().up),
        });
    }
}

/// - Up vector: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.upvector?view=openxml-3.0.1
/// - NormalVector: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.normal?view=openxml-3.0.1
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Vector {
    /// Distance along X-axis in 3D
    pub dx: f64,

    /// Distance along y-axis in 3D
    pub dy: f64,

    /// Distance along z-axis in 3D
    pub dz: f64,
}

impl Vector {
    pub(crate) fn default() -> Self {
        Self {
            dx: 0.0,
            dy: 0.0,
            dz: 0.0,
        }
    }
    pub(crate) fn from_raw(raw: Option<XlsxVector>) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };
        return Self {
            dx: emu_to_pt(raw.dx.unwrap_or(0)),
            dy: emu_to_pt(raw.dy.unwrap_or(0)),
            dz: emu_to_pt(raw.dz.unwrap_or(0)),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.anchor?view=openxml-3.0.1
///
/// Example:
/// ```
/// <anchor x="123" y="23" z="10000"/>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Anchor {
    // x (X-Coordinate in 3D)	X-Coordinate in 3D space.
    pub x: f64,

    // y (Y-Coordinate in 3D)	Y-Coordinate in 3D space.
    pub y: f64,

    // z (Z-Coordinate in 3D)	Z-Coordinate in 3D space.
    pub z: f64,
}

impl Anchor {
    pub(crate) fn default() -> Self {
        Self {
            x: 0.0,
            y: 0.0,
            z: 0.0,
        }
    }
    pub(crate) fn from_raw(raw: Option<XlsxAnchor>) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };
        return Self {
            x: emu_to_pt(raw.x.unwrap_or(0)),
            y: emu_to_pt(raw.y.unwrap_or(0)),
            z: emu_to_pt(raw.z.unwrap_or(0)),
        };
    }
}
