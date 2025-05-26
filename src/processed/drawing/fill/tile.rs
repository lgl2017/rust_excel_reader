#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::st_types::{emu_to_pt, st_percentage_to_float};
use crate::{
    processed::drawing::common_types::rectangle_alignment::RectangleAlignmentValues,
    raw::drawing::fill::blip_fill::XlsxTile,
};

use super::tile_flip::TileFlipValues;

/// tile (Tile)
///
/// This element specifies that a BLIP should be tiled to fill the available space.
/// This element defines a "tile" rectangle within the bounding box.
/// The image is encompassed within the tile rectangle, and the tile rectangle is tiled across the bounding box to fill the entire area.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tile?view=openxml-3.0.1
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Tile {
    /// algn (Alignment)
    ///
    /// Specifies where to align the first tile with respect to the shape.
    /// Alignment happens after the scaling, but before the additional offset.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
    pub alignment: RectangleAlignmentValues,

    /// flip (Tile Flipping)
    ///
    /// Specifies the direction(s) in which to flip the source image while tiling.
    /// Images can be flipped horizontally, vertically, or in both directions to fill the entire region.
    ///
    /// * Horizontal
    /// * HorizontalAndVertical
    /// * None
    /// * Vertical
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tileflipvalues?view=openxml-3.0.1
    pub flip: TileFlipValues,

    /// sx (Horizontal Ratio)
    ///
    /// Specifies the amount to horizontally scale the srcRect in percentage
    ///
    pub horizontal_ratio: f64,

    /// sy (Vertical Ratio)
    ///
    /// Specifies the amount to vertically scale the srcRect.
    ///
    /// Example: 0.65 -> 65%
    pub vertical_ratio: f64,

    /// tx (Horizontal Offset)
    ///
    /// Specifies additional horizontal offset after alignment in points
    pub horizontal_offset: f64,

    /// ty (Vertical Offset)
    ///
    /// Specifies additional vertical offset after alignment in points
    pub vertical_offset: f64,
}

impl Tile {
    pub(crate) fn from_raw(raw: XlsxTile) -> Self {
        return Self {
            alignment: RectangleAlignmentValues::from_string(raw.alignment),
            flip: TileFlipValues::from_string(raw.flip),
            horizontal_ratio: st_percentage_to_float(raw.sx.clone().unwrap_or(1 * 1000 * 100)),
            vertical_ratio: st_percentage_to_float(raw.sy.clone().unwrap_or(1 * 1000 * 100)),
            horizontal_offset: emu_to_pt(raw.tx.clone().unwrap_or(0)),
            vertical_offset: emu_to_pt(raw.ty.clone().unwrap_or(0)),
        };
    }
}
