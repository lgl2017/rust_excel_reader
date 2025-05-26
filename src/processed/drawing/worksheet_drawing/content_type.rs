#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    processed::drawing::{
        effect::effect_style::EffectStyle, fill::Fill, graphic::graphic_frame::GraphicFrame,
        image::picture::Picture, line::outline::Outline, shape::connection_shape::ConnectionShape,
        text::font::Font,
    },
    raw::{
        drawing::{
            graphic::graphic_frame::XlsxGraphicFrame,
            image::picture::XlsxPicture,
            scheme::color_scheme::XlsxColorScheme,
            shape::{connection_shape::XlsxConnectionShape, shape_style::XlsxShapeStyle},
            theme::XlsxTheme,
            worksheet_drawing::{
                client_data::XlsxClientData, drawing_content_type::XlsxWorksheetDrawingContentType,
                group_shape::XlsxGroupShape, spreadsheet_extent::XlsxSpreadsheetExtent,
                spreadsheet_position::XlsxSpreadsheetPosition, spreadsheet_shape::XlsxShape,
            },
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{group_shape::GroupShape, spreadsheet_shape::Shape};

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum DrawingContentType {
    Picture(Picture),
    Shape(Shape),
    GroupShape(GroupShape),
    ConnectionShape(ConnectionShape),
    GraphicFrame(GraphicFrame),
}

impl DrawingContentType {
    pub(crate) fn from_raw(
        raw: Option<XlsxWorksheetDrawingContentType>,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };
        match raw {
            XlsxWorksheetDrawingContentType::ConnectionShape(connection_shape) => {
                return Some(Self::from_connection_shape(
                    connection_shape,
                    extent,
                    position,
                    client_data,
                    drawing_relationship,
                    defined_names,
                    color_scheme,
                    theme,
                ));
            }
            XlsxWorksheetDrawingContentType::GroupShape(group) => {
                return Some(Self::from_group_shape(
                    group,
                    extent,
                    position,
                    client_data,
                    drawing_relationship,
                    image_bytes,
                    defined_names,
                    color_scheme,
                    theme,
                ))
            }
            XlsxWorksheetDrawingContentType::Shape(shape) => {
                return Some(Self::from_shape(
                    shape,
                    extent,
                    position,
                    client_data,
                    parent_group_fill,
                    drawing_relationship,
                    image_bytes,
                    defined_names,
                    color_scheme,
                    theme,
                ));
            }
            XlsxWorksheetDrawingContentType::Picture(picture) => {
                return Self::from_picture(
                    picture,
                    extent,
                    position,
                    client_data,
                    parent_group_fill,
                    drawing_relationship,
                    image_bytes,
                    defined_names,
                    color_scheme,
                    theme,
                );
            }
            XlsxWorksheetDrawingContentType::GraphicFrame(graphic_frame) => {
                return Some(Self::from_graphic_frame(
                    graphic_frame,
                    extent,
                    position,
                    client_data,
                    drawing_relationship,
                    defined_names,
                ));
            }
        };
    }

    fn from_graphic_frame(
        graphic: XlsxGraphicFrame,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let graphic_frame = GraphicFrame::from_spreadsheet_graphic_frame(
            graphic,
            extent,
            position,
            client_data.clone(),
            drawing_relationship.clone(),
            defined_names.clone(),
        );

        return Self::GraphicFrame(graphic_frame);
    }

    fn from_group_shape(
        shape: XlsxGroupShape,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Self {
        let shape = GroupShape::from_raw(
            shape,
            extent,
            position,
            client_data.clone(),
            drawing_relationship.clone(),
            image_bytes.clone(),
            defined_names.clone(),
            color_scheme.clone(),
            theme.clone(),
        );

        return Self::GroupShape(shape);
    }

    fn from_shape(
        shape: XlsxShape,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Self {
        let references = Self::get_references(
            shape.shape_style.clone(),
            color_scheme.clone(),
            theme.clone(),
        );
        let shape = Shape::from_raw(
            shape,
            extent,
            position,
            client_data.clone(),
            parent_group_fill.clone(),
            drawing_relationship.clone(),
            image_bytes.clone(),
            defined_names.clone(),
            color_scheme.clone(),
            references.0,
            references.1,
            references.2,
            references.3,
        );

        return Self::Shape(shape);
    }

    fn from_connection_shape(
        shape: XlsxConnectionShape,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Self {
        let references = Self::get_references(
            shape.shape_style.clone(),
            color_scheme.clone(),
            theme.clone(),
        );
        let shape = ConnectionShape::from_spreadsheet_connection_shape(
            shape,
            extent,
            position,
            client_data.clone(),
            drawing_relationship.clone(),
            defined_names.clone(),
            color_scheme.clone(),
            references.0,
            references.2,
        );

        return Self::ConnectionShape(shape);
    }

    fn from_picture(
        picture: XlsxPicture,
        extent: Option<XlsxSpreadsheetExtent>,
        position: Option<XlsxSpreadsheetPosition>,
        client_data: Option<XlsxClientData>,
        parent_group_fill: Option<Fill>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        defined_names: XlsxDefinedNames,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> Option<Self> {
        let references = Self::get_references(
            picture.shape_style.clone(),
            color_scheme.clone(),
            theme.clone(),
        );
        let Some(picture) = Picture::from_spreadsheet_picture(
            picture,
            extent,
            position,
            client_data.clone(),
            parent_group_fill.clone(),
            drawing_relationship.clone(),
            image_bytes.clone(),
            defined_names.clone(),
            color_scheme.clone(),
            references.0,
            references.1,
            references.2,
        ) else {
            return None;
        };
        return Some(Self::Picture(picture));
    }

    /// get references from theme elemet defined in `XlsxShapeStyle`
    fn get_references(
        raw: Option<XlsxShapeStyle>,
        color_scheme: Option<XlsxColorScheme>,
        theme: Option<Box<XlsxTheme>>,
    ) -> (
        Option<Outline>,
        Option<Fill>,
        Option<EffectStyle>,
        (Option<HexColor>, Option<Font>), // (font color ref, font typeface ref)
    ) {
        let Some(theme) = theme else {
            return (None, None, None, (None, None));
        };

        let Some(raw) = raw else {
            return (None, None, None, (None, None));
        };

        let line_ref = raw.clone().line_reference;
        let raw_line = theme.get_line_from_ref(line_ref.clone());
        let mut ref_color: Option<HexColor> = None;
        if let Some(r) = line_ref.clone() {
            if let Some(color) = r.color {
                ref_color = color.to_hex(color_scheme.clone(), None);
            }
        };
        let line = Outline::from_raw(raw_line, color_scheme.clone(), ref_color);

        let fill_ref = raw.clone().fill_reference;
        let raw_fill = theme.get_fill_from_ref(fill_ref.clone());
        let mut ref_color: Option<HexColor> = None;
        if let Some(r) = fill_ref.clone() {
            if let Some(color) = r.color {
                ref_color = color.to_hex(color_scheme.clone(), None);
            }
        };
        let fill = Fill::from_raw(
            raw_fill,
            None,
            vec![],
            BTreeMap::new(),
            color_scheme.clone(),
            ref_color,
        );

        let effect_ref = raw.clone().effect_reference;
        let raw_effect = theme.get_effect_from_ref(effect_ref.clone());
        let mut ref_color: Option<HexColor> = None;
        if let Some(r) = effect_ref.clone() {
            if let Some(color) = r.color {
                ref_color = color.to_hex(color_scheme.clone(), None);
            }
        };
        let effect = EffectStyle::from_raw(
            raw_effect,
            vec![],
            BTreeMap::new(),
            color_scheme.clone(),
            ref_color,
        );

        let font_ref = raw.clone().font_reference;

        let font_color = if let Some(r) = font_ref.clone() {
            if let Some(color) = r.color {
                color.to_hex(color_scheme.clone(), None)
            } else {
                None
            }
        } else {
            None
        };
        let raw_font = theme.get_font_from_ref(font_ref.clone());
        let font = Font::from_raw(raw_font.clone());

        return (line, fill, effect, (font_color, font));
    }
}
