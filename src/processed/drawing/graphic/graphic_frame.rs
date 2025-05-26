#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        shape::shape_properties::ShapeProperties,
        worksheet_drawing::non_visual_properties::NonVisualDrawingProperty,
    },
    raw::{
        drawing::{
            graphic::graphic_frame::XlsxGraphicFrame,
            worksheet_drawing::{
                client_data::XlsxClientData, spreadsheet_extent::XlsxSpreadsheetExtent,
                spreadsheet_position::XlsxSpreadsheetPosition,
            },
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

/// - GraphicFrame: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.graphicframe?view=openxml-3.0.1
/// - SpreadSheet.GraphicFrame: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.graphicframe?view=openxml-3.0.1
///
/// This element describes a single graphical object frame for a spreadsheet which contains a graphical object.
/// The graphic object is provided entirely by the document authors who choose to persist this data within the document.
///
/// Possible objects (Partial):
/// - Charts
/// - Diagram
///
/// Example
/// ```
/// <xdr:graphicFrame macro="">
/// <xdr:nvGraphicFramePr>
/// <xdr:cNvPr id="2" name="Chart 1">
///     <a:extLst>
///         <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
///             <a16:creationId
///                 xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"
///                 id="{9225A8DB-6689-4068-DEE9-84C654EE0AE3}" />
///         </a:ext>
///     </a:extLst>
/// </xdr:cNvPr>
/// <xdr:cNvGraphicFramePr />
/// </xdr:nvGraphicFramePr>
/// <xdr:xfrm>
/// <a:off x="0" y="0" />
/// <a:ext cx="0" cy="0" />
/// </xdr:xfrm>
/// <a:graphic>
/// <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
///     <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
///         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///         r:id="rId1" />
/// </a:graphicData>
/// </a:graphic>
/// </xdr:graphicFrame>
/// ```
///
/// graphicFrame (Graphic Frame)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct GraphicFrame {
    /// Specifies the uniform resource identifier that represents the data stored under this tag.
    ///
    /// The is used to identify the correct 'server' that can process the contents of this tag.
    pub type_uri: String,

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

    /// This element specifies the visual shape properties that can be applied to a shape.
    pub visual_properties: ShapeProperties,

    /// This element specifies all non-visual properties.
    pub non_visual_properties: NonVisualDrawingProperty,
}

impl GraphicFrame {
    /// - SpreadSheet.GraphicFrame: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.graphicframe?view=openxml-3.0.1
    ///
    /// Attributes: Macro and published is included
    pub(crate) fn from_spreadsheet_graphic_frame(
        raw: XlsxGraphicFrame,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut uri: String = String::new();

        if let Some(graphic) = raw.clone().graphic {
            if let Some(data) = graphic.graphic_data {
                uri = data.uri.unwrap_or(String::new())
            }
        }

        return Self {
            type_uri: uri,
            r#macro: Some(raw.clone().r#macro.unwrap_or(String::new())),
            published: Some(raw.clone().published.unwrap_or(false)),
            visual_properties: ShapeProperties::from_graphic_frame_transform(
                raw.clone().transform2d,
                extent,
                position,
            ),
            non_visual_properties:
                NonVisualDrawingProperty::from_non_visual_graphic_frame_properties(
                    raw.clone().non_visual_graphic_frame_properties,
                    client_data.clone(),
                    drawing_relationship,
                    defined_names,
                ),
        };
    }
}
