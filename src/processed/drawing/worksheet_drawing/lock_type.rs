#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::non_visual_properties::{
    connection_shape_locks::XlsxConnectionShapeLocks, graphic_frame_locks::XlsxGraphicFrameLocks,
    group_shape_locks::XlsxGroupShapeLocks, picture_locks::XlsxPictureLocks,
    shape_locks::XlsxShapeLocks,
};

/// Lock types:
/// * ConnectionShapeLocks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionshapelocks?view=openxml-3.0.1
/// * ContentPartLocks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.drawing.contentpartlocks?view=openxml-3.0.1
/// * Graphic Frame Locks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.graphicframelocks?view=openxml-3.0.1
/// * Group Shape Locks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.groupshapelocks?view=openxml-3.0.1
/// * Picture Locks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picturelocks?view=openxml-3.0.1
/// * Shape Locks: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapelocks?view=openxml-3.0.1
///
/// Possible Values:
/// * NoAdjustHandles
/// * NoChangeArrowheads
/// * NoChangeAspectRatio
/// * NoChangeShapeType
/// * NoDrilldown
/// * NoEditPoints
/// * NoGrouping
/// * NoMove
/// * NoResize
/// * NoSelect
/// * NoRotation
/// * NoTextEdit
/// * NoUngrouping
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum LockTypeValues {
    NoAdjustHandles,
    NoChangeArrowheads,
    NoChangeAspectRatio,
    NoChangeShapeType,
    NoDrilldown,
    NoEditPoints,
    NoGrouping,
    NoMove,
    NoResize,
    NoSelect,
    NoRotation,
    NoTextEdit,
    NoUngrouping,
}

impl LockTypeValues {
    pub(crate) fn from_picture_locks(raw: Option<XlsxPictureLocks>) -> Vec<Self> {
        let Some(raw) = raw else { return vec![] };
        let mut locks: Vec<Self> = vec![];
        if raw.no_adjust_handles == Some(true) {
            locks.push(Self::NoAdjustHandles);
        }
        if raw.no_change_arrowheads == Some(true) {
            locks.push(Self::NoChangeArrowheads);
        }
        if raw.no_aspect_ratio_change == Some(true) {
            locks.push(Self::NoChangeAspectRatio);
        }

        if raw.no_change_shape_type == Some(true) {
            locks.push(Self::NoChangeShapeType);
        }

        if raw.no_edit_points == Some(true) {
            locks.push(Self::NoEditPoints);
        }

        if raw.no_grouping == Some(true) {
            locks.push(Self::NoGrouping);
        }

        if raw.no_move == Some(true) {
            locks.push(Self::NoMove);
        }

        if raw.no_resize == Some(true) {
            locks.push(Self::NoResize);
        }

        if raw.no_select == Some(true) {
            locks.push(Self::NoSelect);
        }

        if raw.no_rotation == Some(true) {
            locks.push(Self::NoRotation);
        }

        return locks;
    }

    pub(crate) fn from_shape_locks(raw: Option<XlsxShapeLocks>) -> Vec<Self> {
        let Some(raw) = raw else { return vec![] };
        let mut locks: Vec<Self> = vec![];
        if raw.no_adjust_handles == Some(true) {
            locks.push(Self::NoAdjustHandles);
        }
        if raw.no_change_arrowheads == Some(true) {
            locks.push(Self::NoChangeArrowheads);
        }
        if raw.no_aspect_ratio_change == Some(true) {
            locks.push(Self::NoChangeAspectRatio);
        }

        if raw.no_change_shape_type == Some(true) {
            locks.push(Self::NoChangeShapeType);
        }

        if raw.no_edit_points == Some(true) {
            locks.push(Self::NoEditPoints);
        }

        if raw.no_grouping == Some(true) {
            locks.push(Self::NoGrouping);
        }

        if raw.no_move == Some(true) {
            locks.push(Self::NoMove);
        }

        if raw.no_resize == Some(true) {
            locks.push(Self::NoResize);
        }

        if raw.no_select == Some(true) {
            locks.push(Self::NoSelect);
        }

        if raw.no_rotation == Some(true) {
            locks.push(Self::NoRotation);
        }

        if raw.no_text_edit == Some(true) {
            locks.push(Self::NoTextEdit);
        }

        return locks;
    }

    pub(crate) fn from_group_shape_locks(raw: Option<XlsxGroupShapeLocks>) -> Vec<Self> {
        let Some(raw) = raw else { return vec![] };
        let mut locks: Vec<Self> = vec![];

        if raw.no_aspect_ratio_change == Some(true) {
            locks.push(Self::NoChangeAspectRatio);
        }

        if raw.no_grouping == Some(true) {
            locks.push(Self::NoGrouping);
        }

        if raw.no_move == Some(true) {
            locks.push(Self::NoMove);
        }

        if raw.no_resize == Some(true) {
            locks.push(Self::NoResize);
        }

        if raw.no_select == Some(true) {
            locks.push(Self::NoSelect);
        }

        if raw.no_rotation == Some(true) {
            locks.push(Self::NoRotation);
        }

        if raw.no_ungrouping == Some(true) {
            locks.push(Self::NoUngrouping);
        }

        return locks;
    }

    pub(crate) fn from_connection_shape_locks(raw: Option<XlsxConnectionShapeLocks>) -> Vec<Self> {
        let Some(raw) = raw else { return vec![] };
        let mut locks: Vec<Self> = vec![];

        if raw.no_adjust_handles == Some(true) {
            locks.push(Self::NoAdjustHandles);
        }
        if raw.no_change_arrowheads == Some(true) {
            locks.push(Self::NoChangeArrowheads);
        }
        if raw.no_aspect_ratio_change == Some(true) {
            locks.push(Self::NoChangeAspectRatio);
        }

        if raw.no_change_shape_type == Some(true) {
            locks.push(Self::NoChangeShapeType);
        }

        if raw.no_edit_points == Some(true) {
            locks.push(Self::NoEditPoints);
        }

        if raw.no_grouping == Some(true) {
            locks.push(Self::NoGrouping);
        }

        if raw.no_move == Some(true) {
            locks.push(Self::NoMove);
        }

        if raw.no_resize == Some(true) {
            locks.push(Self::NoResize);
        }

        if raw.no_select == Some(true) {
            locks.push(Self::NoSelect);
        }

        if raw.no_rotation == Some(true) {
            locks.push(Self::NoRotation);
        }
        return locks;
    }

    pub(crate) fn from_graphic_frame_locks(raw: Option<XlsxGraphicFrameLocks>) -> Vec<Self> {
        let Some(raw) = raw else { return vec![] };
        let mut locks: Vec<Self> = vec![];
        if raw.no_aspect_ratio_change == Some(true) {
            locks.push(Self::NoChangeAspectRatio);
        }

        if raw.no_drilldown == Some(true) {
            locks.push(Self::NoDrilldown);
        }

        if raw.no_grouping == Some(true) {
            locks.push(Self::NoGrouping);
        }

        if raw.no_move == Some(true) {
            locks.push(Self::NoMove);
        }

        if raw.no_resize == Some(true) {
            locks.push(Self::NoResize);
        }

        if raw.no_select == Some(true) {
            locks.push(Self::NoSelect);
        }

        return locks;
    }
}
