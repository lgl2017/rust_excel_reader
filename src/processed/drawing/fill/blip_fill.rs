#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::image::blip::Blip,
    raw::drawing::{fill::blip_fill::XlsxBlipFill, scheme::color_scheme::XlsxColorScheme},
};

use super::{fill_rectangle::FillRectangle, tile::Tile};

/// BlipFill (Picture Fill): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blipfill?view=openxml-3.0.1
///
/// specifies the type of picture fill that a picture object has.
///
/// Example:
/// ```
/// <p:blipFill>
///     <a:blip r:embed="rId2"/>
///     <a:stretch>
///         <a:fillRect b="10000" r="25000"/>
///     </a:stretch>
/// </p:blipFill>
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct BlipFill {
    /// blip (Blip)
    ///
    /// This element specifies the existence of an image (binary large image or picture) and contains a reference to the image data.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip?view=openxml-3.0.1
    pub blip: Blip,

    /// srcRect (Source Rectangle)
    ///
    /// This element specifies the portion of the blip used for the fill.
    /// Each edge of the source rectangle is defined by a percentage offset from the corresponding edge of the bounding box. A positive percentage specifies an inset, while a negative percentage specifies an outset.
    /// Note: For example, a left offset of 25% specifies that the left edge of the source rectangle is located to the right of the bounding box's left edge by an amount equal to 25% of the bounding box's width.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.sourcerectangle?view=openxml-3.0.1
    pub source_rect: FillRectangle,

    /// fill type of the blip
    ///
    /// * Stretch: specifies that a blip should be stretched to fill the target rectangle
    /// * Tile: specifies that a BLIP should be tiled to fill the available space
    pub fill_type: BlipFillType,

    ///  dpi (DPI Setting)
    ///
    /// Specifies the DPI (dots per inch) used to calculate the size of the blip.
    /// If not present or zero, the DPI in the blip is used.
    pub dpi: u64,

    /// rotWithShape (Rotate With Shape)
    ///
    /// Specifies that the fill should rotate with the shape.
    pub rotate_with_shape: bool,
}

impl BlipFill {
    pub(crate) fn from_raw(
        raw: Option<XlsxBlipFill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };
        let Some(blip) = Blip::from_raw(
            raw.clone().blip,
            drawing_relationship,
            image_bytes,
            color_scheme.clone(),
            ref_color.clone(),
        ) else {
            return None;
        };
        return Some(Self {
            blip,
            source_rect: FillRectangle::from_raw(raw.clone().source_rect),
            fill_type: BlipFillType::from_raw(raw.clone()),
            dpi: raw.clone().dpi.unwrap_or(0),
            rotate_with_shape: raw.clone().rot_with_shape.unwrap_or(false),
        });
    }
}

/// * Stretch: specifies that a blip should be stretched to fill the target rectangle
/// * Tile: specifies that a BLIP should be tiled to fill the available space
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BlipFillType {
    /// specifies that a blip should be stretched to fill the target rectangle
    Stretch(FillRectangle),
    /// specifies that a BLIP should be tiled to fill the available space
    Tile(Tile),
}

impl BlipFillType {
    pub(crate) fn from_raw(raw: XlsxBlipFill) -> Self {
        if let Some(stretch) = raw.stretch {
            return Self::Stretch(FillRectangle::from_raw(stretch.fill_rectangle));
        }
        if let Some(raw_tile) = raw.tile {
            return Self::Tile(Tile::from_raw(raw_tile));
        };
        return Self::Stretch(FillRectangle::default());
    }
}
