#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use crate::{
    common_types::{HexColor, Text},
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::{scheme::color_scheme::XlsxColorScheme, text::text_field::XlsxTextField},
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{
    font::Font, paragraph::paragraph_properties::ParagraphProperties,
    text_run_properties::TextRunProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.field?view=openxml-3.0.1
///
/// This element specifies a text field which contains generated text that the application should update periodically.
///
/// Each piece of text when it is generated is given a unique identification number that is used to refer to a specific field.
/// At the time of creation the text field indicates the kind of text that should be used to update this field.
/// This update type is used so that all applications that did not create this text field can still know what kind of text it should be updated with. Thus the new application can then attach an update type to the text field id for continual updating.
///
/// Example:
/// ```
/// <a:fld id="{424CEEAC-8F67-4238-9622-1B74DC6E8318}" type="slidenum">
///     <a:rPr lang="en-US" smtClean="0"/>
///     <a:pPr/>
///     <a:t>3</a:t>
/// </a:fld>
/// ```

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct TextField {
    /// t (Text String)
    pub text: Text,

    /// rPr (Text Run Properties)
    pub run_properties: TextRunProperties,

    /// Specifies the type of text that should be used to update this text field.
    ///
    /// This is used to inform the rendering application what text it should use to update this text field.
    /// There are no specific syntax restrictions placed on this attribute.
    /// The generating application can use it to represent any text that should be updated before rendering the presentation.
    ///
    /// Reserved Values:
    /// - slidenum: presentation slide number
    /// - datetime: default date time format for the rendering application
    /// - datetime1: MM/DD/YYYY date time format [Example: end example]
    /// - datetime2: Day, Month DD, YYYY date time format [Example: Friday, end example]
    /// - datetime3: DD Month YYYY date time format [Example: 12 October 2007 end example]
    /// - datetime4: Month DD, YYYY date time format [Example: end example]
    /// - datetime5: DD-Mon-YY date time format [Example: 12-Oct-07 end example]
    /// - datetime6: Month YY date time format [Example: October 07 end example]
    /// - datetime7: Mon-YY date time format [Example: Oct-07 end example]
    /// - datetime8: MM/DD/YYYY hh:mm AM/PM date time format [Example: end example]
    /// - datetime9:  MM/DD/YYYY hh:mm:ss AM/PM date time format [Example: 4:28:34 PM end example]
    /// - datetime10: hh:mm date time format [Example: 16:28 end example]
    /// - datetime11: hh:mm:ss date time format [Example: 16:28:34 end example]
    /// - datetime12: hh:mm AM/PM date time format [Example: end example]
    /// - datetime13: hh:mm:ss: AM/PM date time format [Example: 4:28:34 PM end example]
    ///
    /// For type `TxLink`, caculation (cell) reference is defined in the sp (Shpae) textlink property.
    pub r#type: String,

    /// paragraph properties
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub paragraph_properties: Option<ParagraphProperties>,
}

impl TextField {
    pub(crate) fn from_raw(
        raw: XlsxTextField,
        default_properties: Option<TextRunProperties>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        return Self {
            run_properties: TextRunProperties::from_raw(
                raw.clone().run_properties,
                default_properties.clone(),
                drawing_relationship.clone(),
                defined_names.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_font.clone(),
                font_ref_color.clone(),
            ),
            text: raw.clone().text.unwrap_or(String::new()),
            r#type: raw.clone().r#type.unwrap_or(String::new()),
            paragraph_properties: if let Some(pr) = raw.clone().paragraph_properties {
                Some(ParagraphProperties::from_raw(
                    Some(*pr),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                ))
            } else {
                None
            },
        };
    }
}
