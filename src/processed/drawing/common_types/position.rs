#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::shape::{
    path::path_point::XlsxPoint, position::XlsxPosition, shape_guide::XlsxShapeGuide,
};

use super::adjust_coordinate::AdjustCoordinate;

/// specifies the position of a drawing element within a spreadsheet
#[derive(Debug, PartialEq, PartialOrd, Clone)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Position {
    pub x: AdjustCoordinate,
    pub y: AdjustCoordinate,
}

impl Position {
    pub(crate) fn default() -> Self {
        Self {
            x: AdjustCoordinate::default(),
            y: AdjustCoordinate::default(),
        }
    }

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.position?view=openxml-3.0.1
    pub(crate) fn from_position(
        position: Option<XlsxPosition>,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let Some(position) = position else {
            return Self::default();
        };
        return Self {
            x: AdjustCoordinate::from_raw(position.x, guide_list.clone()),
            y: AdjustCoordinate::from_raw(position.y, guide_list.clone()),
        };
    }

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.point?view=openxml-3.0.1
    pub(crate) fn from_point(
        point: Option<XlsxPoint>,
        guide_list: Option<Vec<XlsxShapeGuide>>,
    ) -> Self {
        let Some(point) = point else {
            return Self::default();
        };
        return Self {
            x: AdjustCoordinate::from_raw(point.x, guide_list.clone()),
            y: AdjustCoordinate::from_raw(point.y, guide_list.clone()),
        };
    }
}
