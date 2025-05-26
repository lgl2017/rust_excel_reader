#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::worksheet_drawing::{
    one_cell_anchor::XlsxOneCellAnchor, two_cell_anchor::XlsxTwoCellAnchor,
};

use super::cell_marker::CellMarker;

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum DrawingAnchorType {
    /// Move With Cells but Do Not Resize.
    OneCellAnchor(CellMarker),

    /// Move and Resize With Anchor Cells.
    TwoCellAnchor(CellMarker, CellMarker),

    /// Do Not Move or Resize With Underlying Rows/Columns.
    AbsoluteAnchor,
}

impl DrawingAnchorType {
    pub(crate) fn from_two_cell_anchor(raw: XlsxTwoCellAnchor) -> Self {
        // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.editasvalues?view=openxml-3.0.1
        let edit_as = raw.edit_as.unwrap_or("twoCell".to_string());

        return match edit_as.as_ref() {
            "oneCell" => Self::OneCellAnchor(CellMarker::from_raw(raw.from)),
            "absolute" => Self::AbsoluteAnchor,
            _ => Self::TwoCellAnchor(CellMarker::from_raw(raw.from), CellMarker::from_raw(raw.to)),
        };
    }

    pub(crate) fn from_one_cell_anchor(raw: XlsxOneCellAnchor) -> Self {
        return Self::OneCellAnchor(CellMarker::from_raw(raw.from));
    }
}
