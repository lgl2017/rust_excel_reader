pub mod anchor_type;
pub mod cell_marker;
pub mod content_type;
pub mod group_shape;
pub mod lock_type;
pub mod non_visual_properties;
pub mod spreadsheet_shape;

use anchor_type::DrawingAnchorType;
use content_type::DrawingContentType;

#[cfg(feature = "serde")]
use serde::Serialize;

#[derive(Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct WorksheetDrawing {
    pub anchor: DrawingAnchorType,
    pub content: DrawingContentType,
}
