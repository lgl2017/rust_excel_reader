#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::packaging::relationship::XlsxRelationships;

use super::black_white_mode::BlackWhiteModeValues;
use super::shape_style::ShapeStyle;
use super::transform_2d::Transform2D;

use crate::processed::drawing::{
    common_types::{extent::Extents, offset::Offset},
    effect::effect_style::EffectStyle,
    fill::Fill,
    line::outline::Outline,
};

use crate::raw::drawing::{
    scheme::color_scheme::XlsxColorScheme,
    shape::{
        shape_properties::XlsxShapeProperties, transform_2d::XlsxTransform2D,
        visual_group_shape_properties::XlsxVisualGroupShapeProperties,
    },
    worksheet_drawing::{
        spreadsheet_extent::XlsxSpreadsheetExtent, spreadsheet_position::XlsxSpreadsheetPosition,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeproperties?view=openxml-3.0.1
///
/// This element specifies the visual shape properties that can be applied to a shape.
///
/// Example
/// ```
/// <a:spPr>
///     <a:noFill />
///     <a:ln w="12700" cap="flat">
///         <a:solidFill>
///             <a:srgbClr val="000000" />
///         </a:solidFill>
///         <a:prstDash val="solid" />
///         <a:miter lim="400000" />
///     </a:ln>
///     <a:effectLst />
///     <a:sp3d />
/// </a:spPr>
/// ```
// tag: spPr
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ShapeProperties {
    /// size
    pub size: Extents,

    /// position
    pub position: Offset,

    /// child size (Only applicable to group shape properties: `grpSpPr`)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub child_size: Option<Extents>,

    /// child position (Only applicable to group shape properties: `grpSpPr`)
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub child_position: Option<Offset>,

    /// 2d transformation
    pub transform_2d: Transform2D,

    /// shape style
    pub style: ShapeStyle,

    /// bwMode (Black and White Mode)
    ///
    /// Specifies that the picture should be rendered using only black and white coloring.
    /// That is the coloring information for the picture should be converted to either black or white when rendering the picture.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blackwhitemodevalues?view=openxml-3.0.1
    pub black_white_mode: BlackWhiteModeValues,
}

impl ShapeProperties {
    pub(crate) fn default() -> Self {
        return Self {
            size: Extents::default(),
            position: Offset::default(),
            child_size: None,
            child_position: None,
            transform_2d: Transform2D::default(),
            style: ShapeStyle::default(),
            black_white_mode: BlackWhiteModeValues::default(),
        };
    }

    pub(crate) fn from_graphic_frame_transform(
        transform2d: Option<XlsxTransform2D>,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
    ) -> Self {
        let extent = if let Some(transform) = transform2d.clone() {
            Extents::from_transform_2d(Some(transform))
        } else {
            Extents::from_raw(extent.clone())
        };

        let position = if let Some(transform) = transform2d.clone() {
            Offset::from_transform_2d(Some(transform))
        } else {
            Offset::from_spreadsheet_position(position.clone())
        };

        return Self {
            size: extent,
            position,
            child_size: None,
            child_position: None,
            transform_2d: Transform2D::from_transform_2d(transform2d.clone()),
            style: ShapeStyle::default(),
            black_white_mode: BlackWhiteModeValues::default(),
        };
    }

    pub(crate) fn from_shape_properties(
        raw: Option<XlsxShapeProperties>,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        line_ref: Option<Outline>,
        fill_ref: Option<Fill>,
        effect_ref: Option<EffectStyle>,
    ) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        let extent = if let Some(transform) = raw.clone().transform2d {
            Extents::from_transform_2d(Some(transform))
        } else {
            Extents::from_raw(extent.clone())
        };

        let position = if let Some(transform) = raw.clone().transform2d {
            Offset::from_transform_2d(Some(transform))
        } else {
            Offset::from_spreadsheet_position(position.clone())
        };

        return Self {
            size: extent,
            position,
            child_size: None,
            child_position: None,
            transform_2d: Transform2D::from_transform_2d(raw.clone().transform2d),
            style: ShapeStyle::from_shape_properties(
                raw.clone(),
                parent_group_fill.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                line_ref.clone(),
                fill_ref.clone(),
                effect_ref.clone(),
            ),
            black_white_mode: BlackWhiteModeValues::from_string(raw.clone().black_white_mode),
        };
    }

    // group shape does not inherit any parent group fill even when nested in other group shape.
    pub(crate) fn from_group_shape_properties(
        raw: Option<XlsxVisualGroupShapeProperties>,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        let Some(raw) = raw else {
            return Self::default();
        };

        let extent = if let Some(transform) = raw.clone().transform2d {
            Extents::from_transform_group(Some(transform))
        } else {
            Extents::from_raw(extent.clone())
        };

        let position = if let Some(transform) = raw.clone().transform2d {
            Offset::from_transform_group(Some(transform))
        } else {
            Offset::from_spreadsheet_position(position.clone())
        };

        return Self {
            size: extent,
            position,
            child_size: Some(Extents::child_extent_from_transform_group(
                raw.clone().transform2d,
            )),
            child_position: Some(Offset::child_pos_from_transform_group(
                raw.clone().transform2d,
            )),
            transform_2d: Transform2D::from_transform_group(raw.clone().transform2d),
            style: ShapeStyle::from_group_shape_properties(
                raw.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
            black_white_mode: BlackWhiteModeValues::from_string(raw.clone().black_white_mode),
        };
    }
}
