use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use super::adjust_value_list::{load_adjust_value_list, AdjustValueList};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetgeometry?view=openxml-3.0.1
///
/// This element specifies when a preset geometric shape should be used instead of a custom geometric shape.
///
/// Example
/// ```
/// <a:prstGeom prst="heart">
/// </a:prstGeom>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct PresetGeometry {
    // Child Elements
    // avLst (List of Shape Adjust Values)	§20.1.9.5
    pub adjust_value_list: Option<AdjustValueList>,

    // attributes
    /// Preset Shape.
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapetypevalues?view=openxml-3.0.1
    // tag: prst
    pub preset: Option<String>,
}

impl PresetGeometry {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut geom = Self {
            adjust_value_list: None,
            preset: None,
        };
        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"prst" => {
                            geom.preset = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"avLst" => {
                    geom.adjust_value_list = Some(load_adjust_value_list(reader)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"prstGeom" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(geom)
    }
}
