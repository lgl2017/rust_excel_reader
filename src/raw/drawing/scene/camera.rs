use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::helper::string_to_int;

use super::rotation::XlsxRotation;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.camera?view=openxml-3.0.1
/// Example:
/// ```
/// <a:camera prst="orthographicFront">
///     <a:rot lat="19902513" lon="17826689" rev="1362739"/>
/// </a:camera>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCamera {
    // Children
    pub rot: Option<XlsxRotation>,

    // attibutes
    /// Preset Camera Type
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetcameravalues?view=openxml-3.0.1
    pub prst: Option<String>,

    /// Field of view
    pub fov: Option<i64>,

    /// Zoom
    pub zoom: Option<i64>,
}

impl XlsxCamera {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut camera = Self {
            rot: None,
            prst: None,
            fov: None,
            zoom: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"prst" => {
                            camera.prst = Some(string_value);
                        }
                        b"fov" => {
                            camera.fov = string_to_int(&string_value);
                        }
                        b"zoom" => {
                            camera.zoom = string_to_int(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rot" => {
                    camera.rot = Some(XlsxRotation::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"camera" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(camera)
    }
}
