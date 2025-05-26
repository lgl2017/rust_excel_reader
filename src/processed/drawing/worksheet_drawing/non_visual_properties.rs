#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    packaging::relationship::XlsxRelationships,
    processed::shared::hyperlink::Hyperlink,
    raw::{
        drawing::{
            non_visual_properties::{
                non_visual_connection_shape_properties::XlsxNonVisualConnectionShapeProperties,
                non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
                non_visual_graphic_frame_properties::XlsxNonVisualGraphicFrameProperties,
                non_visual_group_shape_properties::XlsxNonVisualGroupShapeProperties,
                non_visual_picture_properties::XlsxNonVisualPictureProperties,
                non_visual_shape_properties::XlsxNonVisualShapeProperties,
            },
            worksheet_drawing::client_data::XlsxClientData,
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::lock_type::LockTypeValues;

/// This element specifies all non-visual properties.
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct NonVisualDrawingProperty {
    /// Specifies the id of the shape
    pub id: u64,

    /// Name compatible with Object Model (non-unique).
    pub name: String,

    /// Flag determining to show or hide this element.
    pub hidden: bool,

    /// alt text for the drawing
    pub description: String,

    /// hyperlink on drawing click
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub hyperlink_on_click: Option<Hyperlink>,

    /// hyperlink on drawing hover
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub hyperlink_on_hover: Option<Hyperlink>,

    /// fLocksWithSheet (Locks With Sheet Flag)
    ///
    /// This attribute indicates whether to disable selection on drawing elements when the sheet is protected.
    pub lock_with_sheet: bool,

    /// fPrintsWithSheet (Prints With Sheet Flag)
    ///
    /// This attribute indicates whether to print drawing elements when printing the sheet.
    pub print_with_sheet: bool,

    /// locking properties for a graphic frame
    pub locks: Vec<LockTypeValues>,
}

impl NonVisualDrawingProperty {
    pub(crate) fn default() -> Self {
        Self {
            id: 0,
            name: String::new(),
            locks: vec![],
            hidden: false,
            description: String::new(),
            hyperlink_on_click: None,
            hyperlink_on_hover: None,
            print_with_sheet: true,
            lock_with_sheet: true,
        }
    }

    pub(crate) fn from_non_visual_shape_properties(
        raw: Option<XlsxNonVisualShapeProperties>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut properties = Self::default();

        let Some(raw) = raw else {
            return properties;
        };
        if let Some(shape_properties) = raw.clone().non_visual_shape_drawing_properties {
            properties.locks = LockTypeValues::from_shape_locks(shape_properties.shape_locks);
        };

        properties = Self::add_non_drawing_properties(
            properties,
            raw.clone().non_visual_drawing_properties,
            drawing_relationship,
            defined_names,
        );
        return Self::add_client_data_properties(properties, client_data);
    }

    pub(crate) fn from_non_visual_group_shape_properties(
        raw: Option<XlsxNonVisualGroupShapeProperties>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut properties = Self::default();

        let Some(raw) = raw else {
            return properties;
        };
        if let Some(shape_properties) = raw.clone().non_visual_group_shape_drawing_properties {
            properties.locks =
                LockTypeValues::from_group_shape_locks(shape_properties.group_shape_locks);
        };

        properties = Self::add_non_drawing_properties(
            properties,
            raw.clone().non_visual_drawing_properties,
            drawing_relationship,
            defined_names,
        );
        return Self::add_client_data_properties(properties, client_data);
    }

    pub(crate) fn from_non_visual_picture_properties(
        raw: Option<XlsxNonVisualPictureProperties>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut properties = Self::default();

        let Some(raw) = raw else {
            return properties;
        };
        if let Some(picture_properties) = raw.clone().non_visual_picture_drawing_properties {
            properties.locks = LockTypeValues::from_picture_locks(picture_properties.picture_locks);
        };

        properties = Self::add_non_drawing_properties(
            properties,
            raw.clone().non_visual_drawing_properties,
            drawing_relationship,
            defined_names,
        );
        return Self::add_client_data_properties(properties, client_data);
    }

    pub(crate) fn from_non_visual_connection_shape_properties(
        raw: Option<XlsxNonVisualConnectionShapeProperties>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut properties = Self::default();

        let Some(raw) = raw else {
            return properties;
        };
        if let Some(shape_properties) = raw.clone().non_visual_connector_shape_drawing_properties {
            properties.locks = LockTypeValues::from_connection_shape_locks(
                shape_properties.connection_shape_locks,
            );
        };

        properties = Self::add_non_drawing_properties(
            properties,
            raw.clone().non_visual_drawing_properties,
            drawing_relationship,
            defined_names,
        );
        return Self::add_client_data_properties(properties, client_data);
    }

    pub(crate) fn from_non_visual_graphic_frame_properties(
        raw: Option<XlsxNonVisualGraphicFrameProperties>,
        client_data: Option<XlsxClientData>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let mut properties = Self::default();

        let Some(raw) = raw else {
            return properties;
        };
        if let Some(frame_properties) = raw.clone().non_visual_graphic_frame_drawing_properties {
            properties.locks =
                LockTypeValues::from_graphic_frame_locks(frame_properties.graphic_frame_locks);
        };

        properties = Self::add_non_drawing_properties(
            properties,
            raw.clone().non_visual_drawing_properties,
            drawing_relationship,
            defined_names,
        );
        return Self::add_client_data_properties(properties, client_data);
    }

    fn add_non_drawing_properties(
        mut properties: Self,
        drawing_properties: Option<XlsxNonVisualDrawingProperties>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
    ) -> Self {
        let Some(drawing_properties) = drawing_properties else {
            return properties;
        };

        properties.id = drawing_properties.clone().id.unwrap_or(0);
        properties.name = drawing_properties.clone().name.unwrap_or(String::new());
        properties.hidden = drawing_properties.clone().hidden.unwrap_or(false);
        properties.hyperlink_on_click = Hyperlink::from_hlink_event(
            drawing_properties.clone().hlink_click,
            drawing_relationship.clone(),
            defined_names.clone(),
        );
        properties.hyperlink_on_hover = Hyperlink::from_hlink_event(
            drawing_properties.clone().hlink_hover,
            drawing_relationship,
            defined_names,
        );
        properties.description = drawing_properties.description.unwrap_or(String::new());

        return properties;
    }

    fn add_client_data_properties(
        mut properties: Self,
        client_data: Option<XlsxClientData>,
    ) -> Self {
        let Some(client_data) = client_data else {
            return properties;
        };

        properties.lock_with_sheet = client_data.f_locks_with_sheet.unwrap_or(true);
        properties.print_with_sheet = client_data.f_prints_with_sheet.unwrap_or(true);

        return properties;
    }
}
