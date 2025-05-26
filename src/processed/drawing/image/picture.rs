use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        effect::effect_style::EffectStyle,
        fill::{blip_fill::BlipFill, Fill},
        line::outline::Outline,
        shape::{geometry::GeometryTypeValues, shape_properties::ShapeProperties},
        worksheet_drawing::non_visual_properties::NonVisualDrawingProperty,
    },
    raw::{
        drawing::{
            image::picture::XlsxPicture,
            scheme::color_scheme::XlsxColorScheme,
            worksheet_drawing::{
                client_data::XlsxClientData, spreadsheet_extent::XlsxSpreadsheetExtent,
                spreadsheet_position::XlsxSpreadsheetPosition,
            },
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

/// - Picture: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picture?view=openxml-3.0.1
/// - SpreadSheet.Picture: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1
///
/// This element specifies the existence of a picture object within the document.
///
/// Example:
/// ```
/// <p:pic>
///   <p:nvPicPr>
///     <p:cNvPr id="4" name="lake.JPG" descr="Picture of a Lake" />
///     <p:cNvPicPr>
///       <a:picLocks noChangeAspect="1"/>
///     </p:cNvPicPr>
///     <p:nvPr/>
///   </p:nvPicPr>
///   <p:blipFill>
///   …  </p:blipFill>
///   <p:spPr>
///   …  </p:spPr>
/// </p:pic>
/// ```
///
/// pic (Picture)
// Note: shape style not applicable to picture
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Picture {
    /// blipFill (Picture Fill)
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blipfill?view=openxml-3.0.1
    pub blip_fill: BlipFill,

    /// geometry:
    /// * CustomGeometry: This shape consists of a series of lines and curves described within a creation path.
    ///     In addition to this there can also be adjust values, guides, adjust handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
    /// * PresetGeometry: A preset geometric shape should be used instead of a custom geometric shape.
    pub geometry: GeometryTypeValues,

    /// macro: Only Apply to SpreadSheet.ConnectionShape
    ///
    /// Reference to custom function
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub r#macro: Option<String>,

    /// Published:  Only Apply to SpreadSheet.ConnectionShape
    ///
    /// Publish to Server Flag
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub published: Option<bool>,

    /// DrawingContentType::Picture: preferRelativeResize (Relative Resize Preferred)
    ///
    /// Specifies if the user interface should show the resizing of the picture based on the picture's current size or its original size.
    /// If this attribute is set to true, then scaling is relative to the original picture size as opposed to the current picture size.
    pub prefer_relative_resize: bool,

    /// This element specifies the visual shape properties that can be applied to a shape.
    pub visual_properties: ShapeProperties,

    /// This element specifies all non-visual properties for a picture.
    pub non_visual_properties: NonVisualDrawingProperty,
}

impl Picture {
    /// - SpreadSheet.Picture: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.picture?view=openxml-3.0.1
    ///
    /// Attributes: Macro and published is included
    pub(crate) fn from_spreadsheet_picture(
        raw: XlsxPicture,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        line_ref: Option<Outline>,
        fill_ref: Option<Fill>,
        effect_ref: Option<EffectStyle>,
    ) -> Option<Self> {
        let prefer_relative_resize =
            if let Some(non_visual_shape_properties) = raw.non_visual_picture_properties.clone() {
                if let Some(p) = non_visual_shape_properties.non_visual_picture_drawing_properties {
                    p.prefer_relative_resize.unwrap_or(false)
                } else {
                    false
                }
            } else {
                false
            };

        let Some(blip) = BlipFill::from_raw(
            raw.clone().blip_fill,
            drawing_relationship.clone(),
            image_bytes.clone(),
            color_scheme.clone(),
            None,
        ) else {
            return None;
        };

        return Some(Self {
            blip_fill: blip,
            geometry: GeometryTypeValues::from_shape_properties(
                raw.clone().shape_properties.clone(),
            ),
            prefer_relative_resize,
            r#macro: Some(raw.clone().r#macro.unwrap_or(String::new())),
            published: Some(raw.clone().published.unwrap_or(false)),
            non_visual_properties: NonVisualDrawingProperty::from_non_visual_picture_properties(
                raw.non_visual_picture_properties,
                client_data.clone(),
                drawing_relationship.clone(),
                defined_names,
            ),
            visual_properties: ShapeProperties::from_shape_properties(
                raw.shape_properties,
                extent,
                position,
                parent_group_fill.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                line_ref.clone(),
                fill_ref.clone(),
                effect_ref.clone(),
            ),
        });
    }
}
