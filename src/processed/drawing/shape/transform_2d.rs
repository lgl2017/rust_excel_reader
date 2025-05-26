#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    shape::{transform_2d::XlsxTransform2D, transform_group::XlsxTransformGroup},
    st_types::st_angle_to_degree,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transform2d?view=openxml-3.0.1
///
/// This element represents 2-D transforms for ordinary shapes.
///
/// Example
/// ```
/// <a:xfrm>
///   <a:off x="3200400" y="1600200"/>
///   <a:ext cx="1705233" cy="679622"/>
/// </a:xfrm>
/// ```
///
/// xfrm
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Transform2D {
    /// Specifies a horizontal flip.
    /// When true, this attribute defines that the shape is flipped horizontally about the center of its bounding box.
    ///
    /// flipH (Horizontal Flip)
    pub horizontal_flip: bool,

    /// Specifies a vertical flip.
    /// When true, this attribute defines that the group is flipped vertically about the center of its bounding box.
    ///
    /// flipV (Vertical Flip)
    pub vertical_flip: bool,

    /// Specifies the rotation angle of the Graphic Frame.
    ///
    /// rot (Rotation)
    pub rotation: f64,
}

impl Transform2D {
    pub(crate) fn default() -> Self {
        return Self {
            horizontal_flip: false,
            vertical_flip: false,
            rotation: 0.0,
        };
    }
    pub(crate) fn from_transform_2d(raw: Option<XlsxTransform2D>) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        return Self {
            horizontal_flip: raw.horizontal_flip.unwrap_or(false),
            vertical_flip: raw.vertical_flip.unwrap_or(false),
            rotation: st_angle_to_degree(raw.rotation.unwrap_or(0)),
        };
    }

    pub(crate) fn from_transform_group(raw: Option<XlsxTransformGroup>) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        return Self {
            horizontal_flip: raw.horizontal_flip.unwrap_or(false),
            vertical_flip: raw.vertical_flip.unwrap_or(false),
            rotation: st_angle_to_degree(raw.rotation.unwrap_or(0)),
        };
    }
}
