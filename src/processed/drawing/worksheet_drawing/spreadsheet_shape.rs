#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        effect::effect_style::EffectStyle,
        fill::Fill,
        line::outline::Outline,
        shape::{geometry::GeometryTypeValues, shape_properties::ShapeProperties},
        text::{font::Font, shape_text_body::ShapeTextBody},
        worksheet_drawing::non_visual_properties::NonVisualDrawingProperty,
    },
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme,
            worksheet_drawing::{
                client_data::XlsxClientData, spreadsheet_extent::XlsxSpreadsheetExtent,
                spreadsheet_position::XlsxSpreadsheetPosition, spreadsheet_shape::XlsxShape,
            },
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1
///
/// This element specifies the existence of a single shape.
/// A shape can either be a preset or a custom geometry, defined using the DrawingML framework.
/// In addition to a geometry each shape can have both visual and non-visual properties attached.
/// Text and corresponding styling information can also be attached to a shape.
/// This shape is specified along with all other shapes within either the shape tree or group shape elements.
///
/// Example:
/// ```
/// <p:sp macro="" textlink="$D$3">
///   <p:nvSpPr>
///     <p:cNvPr id="2" name="Rectangle 1"/>
///     <p:cNvSpPr>
///       <a:spLocks noGrp="1"/>
///     </p:cNvSpPr>
///   </p:nvSpPr>
/// …</p:sp>
/// ```
///
/// sp (Shape)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Shape {
    /// geometry:
    /// * CustomGeometry: This shape consists of a series of lines and curves described within a creation path.
    ///     In addition to this there can also be adjust values, guides, adjust handles, connection sites and an inscribed rectangle specified for this custom geometric shape.
    /// * PresetGeometry: A preset geometric shape should be used instead of a custom geometric shape.
    pub geometry: GeometryTypeValues,

    /// Text Box
    ///
    /// Specifies that the corresponding shape is a text box and thus should be treated as such by the generating application.
    /// If this attribute is omitted then it is assumed that the corresponding shape is not specifically a text box.
    pub text_box: bool,

    /// Lock Text
    ///
    /// Default to true
    pub lock_text: bool,

    /// Text Link
    ///
    /// Reference used by a fld (TextField) of type `TxLink`.
    ///
    /// Example:
    /// ```
    /// textlink="$D$3"
    /// ```
    pub text_link: String,

    /// This element specifies the existence of a text within a parent shape.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub text: Option<ShapeTextBody>,

    /// macro
    ///
    /// Reference to custom function
    pub r#macro: String,

    /// Published
    ///
    /// Publish to Server Flag
    pub published: bool,

    /// This element specifies the visual shape properties that can be applied to a shape.
    pub visual_properties: ShapeProperties,

    /// This element specifies all non-visual properties.
    pub non_visual_properties: NonVisualDrawingProperty,
}

impl Shape {
    pub(crate) fn from_raw(
        raw: XlsxShape,
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
        font_ref: (Option<HexColor>, Option<Font>), // color, font typefaces
    ) -> Self {
        let text_box =
            if let Some(non_visual_shape_properties) = raw.non_visual_shape_properties.clone() {
                if let Some(p) = non_visual_shape_properties.non_visual_shape_drawing_properties {
                    p.text_box.unwrap_or(false)
                } else {
                    false
                }
            } else {
                false
            };

        return Self {
            geometry: GeometryTypeValues::from_shape_properties(raw.clone().shape_properties),
            text_box,
            lock_text: raw.clone().lock_text.unwrap_or(true),
            text_link: raw.clone().text_link.unwrap_or(String::new()),
            text: ShapeTextBody::from_raw(
                raw.clone().text_body,
                drawing_relationship.clone(),
                defined_names.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                font_ref.1,
                font_ref.0,
            ),
            r#macro: raw.clone().r#macro.unwrap_or(String::new()),
            published: raw.clone().published.unwrap_or(false),
            non_visual_properties: NonVisualDrawingProperty::from_non_visual_shape_properties(
                raw.non_visual_shape_properties,
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
        };
    }
}
