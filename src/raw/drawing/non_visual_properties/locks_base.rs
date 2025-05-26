use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLocksBase {
    // Child Elements
    // extLst (Not Supported)

    // Attributes:
    /// noAdjustHandles:
    ///
    /// A Boolean attribute that specifies that the generating application does not show adjust handles for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_adjust_handles: Option<bool>,

    /// noChangeArrowheads:
    ///
    /// A Boolean attribute that specifies that the generating application does not enable arrowhead changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_change_arrowheads: Option<bool>,

    /// noChangeAspect
    ///
    /// A Boolean attribute that specifies that the generating application does not enable aspect ratio changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_aspect_ratio_change: Option<bool>,

    /// noChangeShapeType:
    ///
    /// A Boolean attribute that specifies that the generating application does not enable shape type changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_change_shape_type: Option<bool>,

    /// noEditPoints:
    ///
    /// A Boolean attribute that specifies that the generating application does not enable shape point changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_edit_points: Option<bool>,

    // Attributes:
    /// noGrp
    ///
    /// A Boolean attribute that specifies that the generating application does not enable shape grouping for the corresponding content part.
    /// That is, it cannot be combined with other shapes to form a group of shapes.
    ///
    /// The default value is FALSE.
    pub no_grouping: Option<bool>,

    /// noMove
    ///
    /// A Boolean attribute that specifies that the generating application does not enable position changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_move: Option<bool>,

    /// noResize:
    ///
    /// A Boolean attribute that specifies that the generating application does not enable size changes for the corresponding content part.
    ///
    /// The default value is FALSE.
    pub no_resize: Option<bool>,

    /// noSelect
    ///
    /// A Boolean attribute that specifies that the generating application does not enable selecting the corresponding content part.
    /// No picture, shapes, or text attached to this content part can be selected if this attribute has been specified.
    ///
    /// The default value is FALSE.
    pub no_select: Option<bool>,

    /// noRot
    ///
    /// A Boolean attribute that specifies that the corresponding content part cannot be rotated.
    ///
    /// The default value is FALSE.
    pub no_rotation: Option<bool>,
}

impl XlsxLocksBase {
    pub(crate) fn load(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
        tag: &[u8],
    ) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            no_grouping: None,
            no_select: None,
            no_rotation: None,
            no_aspect_ratio_change: None,
            no_move: None,
            no_resize: None,
            no_edit_points: None,
            no_adjust_handles: None,
            no_change_arrowheads: None,
            no_change_shape_type: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"noGrp" => {
                            properties.no_grouping = string_to_bool(&string_value);
                        }
                        b"noSelect" => {
                            properties.no_select = string_to_bool(&string_value);
                        }
                        b"noRot" => {
                            properties.no_rotation = string_to_bool(&string_value);
                        }
                        b"noChangeAspect" => {
                            properties.no_aspect_ratio_change = string_to_bool(&string_value);
                        }
                        b"noMove" => {
                            properties.no_move = string_to_bool(&string_value);
                        }
                        b"noResize" => {
                            properties.no_resize = string_to_bool(&string_value);
                        }
                        b"noEditPoints" => {
                            properties.no_edit_points = string_to_bool(&string_value);
                        }
                        b"noAdjustHandles" => {
                            properties.no_adjust_handles = string_to_bool(&string_value);
                        }
                        b"noChangeArrowheads" => {
                            properties.no_change_arrowheads = string_to_bool(&string_value);
                        }
                        b"noChangeShapeType" => {
                            properties.no_change_shape_type = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => {
                    bail!(
                        "unexpected end of file at `{}`.",
                        String::from_utf8(tag.to_vec())?
                    )
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
