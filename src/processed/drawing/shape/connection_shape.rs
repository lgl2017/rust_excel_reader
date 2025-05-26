use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        effect::effect_style::EffectStyle, line::outline::Outline,
        worksheet_drawing::non_visual_properties::NonVisualDrawingProperty,
    },
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme,
            shape::{
                connection_shape::XlsxConnectionShape, end_connection::XlsxEndConnection,
                start_connection::XlsxStartConnection,
            },
            worksheet_drawing::{
                client_data::XlsxClientData, spreadsheet_extent::XlsxSpreadsheetExtent,
                spreadsheet_position::XlsxSpreadsheetPosition,
            },
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{geometry::GeometryTypeValues, shape_properties::ShapeProperties};

/// - ConnectionShape: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionshape?view=openxml-3.0.1
/// - SpreadSheet.ConnectionShape: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1
///
/// This element specifies a connection shape that is used to connect two sp elements.
/// Once a connection is specified using a cxnSp, it is left to the generating application to determine the exact path the connector takes.
/// That is the connector routing algorithm is left up to the generating application as the desired path might be different depending on the specific needs of the application.
///
/// Example:
/// ```
/// <xdr:cxnSp macro="">
///     <xdr:nvCxnSpPr>
///         <xdr:cNvPr id="5" name="Straight Arrow Connector 4">
///         </xdr:cNvPr>
///         <xdr:cNvCxnSpPr>
///             <a:stCxn id="3" idx="1" />
///             <a:endCxn id="2" idx="4" />
///         </xdr:cNvCxnSpPr>
///     </xdr:nvCxnSpPr>
///     <xdr:spPr>
///         <a:xfrm flipH="1" flipV="1">
///             <a:off x="2311398" y="1505943" />
///             <a:ext cx="1333502" cy="741957" />
///         </a:xfrm>
///         <a:prstGeom prst="straightConnector1">
///             <a:avLst />
///         </a:prstGeom>
///         <a:ln>
///             <a:solidFill>
///                 <a:schemeClr val="accent1" />
///             </a:solidFill>
///             <a:bevel />
///             <a:tailEnd type="triangle" w="med" len="lg" />
///         </a:ln>
///     </xdr:spPr>
///     <xdr:style>
///     … </xdr:style>
/// </xdr:cxnSp>
/// ```
///
/// cxnSp (Connection Shape)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ConnectionShape {
    /// geometry:
    /// * CustomGeometry: This shape consists of a series of lines and curves described within a creation path.
    ///     In addition to this there can also be adjust values, guides, adjust handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
    /// * PresetGeometry: A preset geometric shape should be used instead of a custom geometric shape.
    pub geometry: GeometryTypeValues,

    /// Connection Start Point
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub start_connection: Option<ConnectionPoint>,

    /// Connection End Point
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub end_connection: Option<ConnectionPoint>,

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

impl ConnectionShape {
    /// - SpreadSheet.ConnectionShape: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1
    ///
    /// Attributes: Macro and published is included
    // connection shape do not have fills
    pub(crate) fn from_spreadsheet_connection_shape(
        raw: XlsxConnectionShape,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        line_ref: Option<Outline>,
        effect_ref: Option<EffectStyle>,
    ) -> Self {
        let mut start: Option<ConnectionPoint> = None;
        let mut end: Option<ConnectionPoint> = None;

        if let Some(properties) = raw.clone().non_visual_connection_shape_properties {
            if let Some(p) = properties.non_visual_connector_shape_drawing_properties {
                start = ConnectionPoint::from_start(p.start_connection);
                end = ConnectionPoint::from_end(p.end_connection);
            }
        };

        return Self {
            geometry: GeometryTypeValues::from_shape_properties(raw.clone().shape_properties),
            start_connection: start,
            end_connection: end,
            r#macro: Some(raw.clone().r#macro.unwrap_or(String::new())),
            published: Some(raw.clone().published.unwrap_or(false)),
            non_visual_properties:
                NonVisualDrawingProperty::from_non_visual_connection_shape_properties(
                    raw.non_visual_connection_shape_properties,
                    client_data.clone(),
                    drawing_relationship.clone(),
                    defined_names,
                ),
            visual_properties: ShapeProperties::from_shape_properties(
                raw.shape_properties,
                extent,
                position,
                None,
                drawing_relationship.clone(),
                BTreeMap::new(),
                color_scheme.clone(),
                line_ref.clone(),
                None,
                effect_ref.clone(),
            ),
        };
    }
}

/// - [Start Connection](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.startconnection?view=openxml-3.0.1)
/// - [End Conneciton](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.endconnection?view=openxml-3.0.1)
///
/// Example:
/// ```
/// <xdr:cNvCxnSpPr>
///     <a:stCxn id="3" idx="1" />
///     <a:endCxn id="2" idx="4" />
/// </xdr:cNvCxnSpPr
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ConnectionPoint {
    /// id (Identifier)
    ///
    /// Specifies the id of the shape to make the final connection to.
    pub id: u64,

    /// idx (Index)
    ///
    /// Specifies the index into the connection site table of the final connection shape.
    pub index: u64,
}

impl ConnectionPoint {
    pub(crate) fn from_start(raw: Option<XlsxStartConnection>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(id) = raw.id else { return None };
        return Some(Self {
            id,
            index: raw.index.unwrap_or(0),
        });
    }

    pub(crate) fn from_end(raw: Option<XlsxEndConnection>) -> Option<Self> {
        let Some(raw) = raw else { return None };

        let Some(id) = raw.id else { return None };
        return Some(Self {
            id,
            index: raw.index.unwrap_or(0),
        });
    }
}
