#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{
    shape::{
        offset::XlsxOffset, transform_2d::XlsxTransform2D, transform_group::XlsxTransformGroup,
    },
    st_types::emu_to_pt,
    worksheet_drawing::spreadsheet_position::XlsxSpreadsheetPosition,
};

/// specifies the position of a drawing element within a spreadsheet
#[derive(Debug, PartialEq, PartialOrd, Clone)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Offset {
    pub x: f64,
    pub y: f64,
}

impl Offset {
    pub(crate) fn default() -> Self {
        Self { x: 0.0, y: 0.0 }
    }

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.position?view=openxml-3.0.1
    pub(crate) fn from_spreadsheet_position(position: Option<XlsxSpreadsheetPosition>) -> Self {
        let Some(position) = position else {
            return Self::default();
        };
        return Self {
            x: emu_to_pt(position.x.unwrap_or(0)),
            y: emu_to_pt(position.y.unwrap_or(0)),
        };
    }

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.offset?view=openxml-3.0.1
    ///
    ///  Example
    /// ```
    /// <a:off x="3200400" y="1600200"/>
    /// ```
    pub(crate) fn from_offset(offset: Option<XlsxOffset>) -> Self {
        let Some(offset) = offset else {
            return Self::default();
        };
        return Self {
            x: emu_to_pt(offset.x.unwrap_or(0)),
            y: emu_to_pt(offset.y.unwrap_or(0)),
        };
    }

    pub(crate) fn from_transform_2d(transform: Option<XlsxTransform2D>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };
        return Self::from_offset(transform.offset);
    }

    pub(crate) fn from_transform_group(transform: Option<XlsxTransformGroup>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };
        return Self::from_offset(transform.offset);
    }

    pub(crate) fn child_pos_from_transform_group(transform: Option<XlsxTransformGroup>) -> Self {
        let Some(transform) = transform else {
            return Self::default();
        };

        if let Some(child) = transform.child_offset {
            return Self::from_offset(Some(child));
        } else {
            return Self::from_offset(transform.offset);
        }
    }
}
