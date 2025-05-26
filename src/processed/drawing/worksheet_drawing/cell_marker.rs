#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::st_types::emu_to_pt;
use crate::raw::drawing::worksheet_drawing::marker::XlsxMarker;

/// specifies anchoring information for a shape within a spreadsheet.
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, Default, PartialEq, PartialOrd, Copy, Clone)]
pub struct CellMarker {
    /// column (1 based index)
    pub col: u64,

    /// column offset within a cell in point
    pub col_offset: f64,

    /// row (1 based index)
    pub row: u64,

    /// row offset within a cell in point
    pub row_offset: f64,
}

impl CellMarker {
    pub(crate) fn default() -> Self {
        Self {
            col: 1,
            col_offset: 0.0,
            row: 1,
            row_offset: 0.0,
        }
    }
    pub(crate) fn from_raw(marker: Option<XlsxMarker>) -> Self {
        let Some(marker) = marker else {
            return Self::default();
        };
        return Self {
            col: marker.column_id.unwrap_or(0) + 1,
            col_offset: emu_to_pt(marker.column_offset.unwrap_or(0)),
            row: marker.row_id.unwrap_or(0) + 1,
            row_offset: emu_to_pt(marker.row_offset.unwrap_or(0)),
        };
    }
}
