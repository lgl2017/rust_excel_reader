#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::st_types::emu_to_pt;

use crate::raw::drawing::shape::{
    extents::XlsxExtents, transform_2d::XlsxTransform2D, transform_group::XlsxTransformGroup,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.extents?view=openxml-3.0.1
///
/// specifies the size of the bounding box enclosing the referenced object.
///
///  Example
/// ```
/// <a:ext cx="1705233" cy="679622"/>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, Default, PartialEq, PartialOrd, Copy, Clone)]
pub struct Extents {
    // width in points
    pub width: f64,
    // height in points
    pub height: f64,
}

impl Extents {
    pub(crate) fn default() -> Self {
        Self {
            width: 0.0,
            height: 0.0,
        }
    }

    pub(crate) fn from_raw(extent: Option<XlsxExtents>) -> Self {
        let Some(extent) = extent else {
            return Self::default();
        };
        return Self {
            width: emu_to_pt(extent.cx.unwrap_or(0)),
            height: emu_to_pt(extent.cy.unwrap_or(0)),
        };
    }

    pub(crate) fn from_transform_2d(transform: Option<XlsxTransform2D>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };
        return Self::from_raw(transform.extents);
    }

    pub(crate) fn from_transform_group(transform: Option<XlsxTransformGroup>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };
        return Self::from_raw(transform.extents);
    }

    pub(crate) fn child_extent_from_transform_group(transform: Option<XlsxTransformGroup>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };
        if let Some(child) = transform.child_extents {
            return Self::from_raw(Some(child));
        } else {
            return Self::from_raw(transform.extents);
        }
    }
}
