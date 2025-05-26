#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::raw::drawing::{
    scheme::color_scheme::XlsxColorScheme,
    theme::XlsxTheme,
    worksheet_drawing::{
        client_data::XlsxClientData, group_shape::XlsxGroupShape,
        spreadsheet_extent::XlsxSpreadsheetExtent, spreadsheet_position::XlsxSpreadsheetPosition,
    },
};
use crate::raw::spreadsheet::workbook::defined_name::XlsxDefinedNames;
use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        fill::Fill,
        shape::shape_properties::ShapeProperties,
        worksheet_drawing::{
            content_type::DrawingContentType, non_visual_properties::NonVisualDrawingProperty,
        },
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.groupshape?view=openxml-3.0.1
///
/// This element specifies a group shape that represents many shapes grouped together.
/// This shape is to be treated just as if it were a regular shape but instead of being described by a single geometry it is made up of all the shape geometries encompassed within it.
/// Within a group shape each of the shapes that make up the group are specified just as they normally would.
/// The idea behind grouping elements however is that a single transform can apply to many shapes at the same time.
///
/// Example:
///
/// ```
/// <p:grpSp>
///   <p:nvGrpSpPr>
///     <p:cNvPr id="10" name="Group 9"/>
///     <p:cNvGrpSpPr/>
///     <p:nvPr/>
///   </p:nvGrpSpPr>
///   <p:grpSpPr>
///     <a:xfrm>
///       <a:off x="838200" y="990600"/>
///       <a:ext cx="2426208" cy="978408"/>
///       <a:chOff x="838200" y="990600"/>
///       <a:chExt cx="2426208" cy="978408"/>
///     </a:xfrm>
///   </p:grpSpPr>
///   <p:sp>
///   …  </p:sp>
///   <p:sp>
///   …  </p:sp>
///   <p:sp>
///   …  </p:sp>
/// </p:grpSp>
/// ```
///
/// grpSp (Group shape)

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct GroupShape {
    /// grouped drawing contens
    pub contents: Vec<DrawingContentType>,

    /// This element specifies the visual shape properties that can be applied to a shape.
    pub visual_properties: ShapeProperties,

    /// This element specifies all non-visual properties.
    pub non_visual_properties: NonVisualDrawingProperty,
}

impl GroupShape {
    pub(crate) fn from_raw(
        raw: XlsxGroupShape,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Self {
        let fill = if let Some(visual_group_shape_properties) =
            raw.clone().visual_group_shape_properties
        {
            Fill::from_group_shape_properties(
                visual_group_shape_properties,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            )
        } else {
            None
        };

        let contents: Vec<DrawingContentType> = raw
            .clone()
            .drawing_contents
            .unwrap_or(Box::new(vec![]))
            .into_iter()
            .map(|c| {
                DrawingContentType::from_raw(
                    Some(c),
                    // grouped children don't use the main extent and position from the drawing
                    None,
                    None,
                    client_data.clone(),
                    fill.clone(),
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    defined_names.clone(),
                    color_scheme.clone(),
                    theme.clone(),
                )
            })
            .filter(|c| c.is_some())
            .map(|c| c.unwrap())
            .collect();

        return Self {
            contents,
            non_visual_properties: NonVisualDrawingProperty::from_non_visual_group_shape_properties(
                raw.non_visual_group_shape_properties,
                client_data.clone(),
                drawing_relationship.clone(),
                defined_names,
            ),
            visual_properties: ShapeProperties::from_group_shape_properties(
                raw.visual_group_shape_properties,
                extent,
                position,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
        };
    }
}
