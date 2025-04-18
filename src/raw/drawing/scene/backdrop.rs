use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backdrop?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBackDrop {
    // extLst Not supported

    // Child Elements
    // anchor
    pub anchor: Option<XlsxAnchor>,
    // norm
    norm: Option<XlsxNormalVector>,
    // up
    up: Option<XlsxUpVector>,
}

impl XlsxBackDrop {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut list = Self {
            anchor: None,
            norm: None,
            up: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"anchor" => {
                    list.anchor = Some(XlsxAnchor::load(e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"norm" => {
                    list.norm = Some(XlsxNormalVector::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"up" => {
                    list.up = Some(XlsxUpVector::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"backdrop" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(list)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.normal?view=openxml-3.0.1
///
/// Example:
/// ```
/// <norm dx="123" dy="23" dz="10000"/>
/// ```
pub type XlsxNormalVector = XlsxVector;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.upvector?view=openxml-3.0.1
pub type XlsxUpVector = XlsxVector;

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxVector {
    // attributes
    /// Distance along X-axis in 3D
    dx: Option<i64>,

    /// Distance along y-axis in 3D
    dy: Option<i64>,

    /// Distance along z-axis in 3D
    dz: Option<i64>,
}

impl XlsxVector {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut vector = Self {
            dx: None,
            dy: None,
            dz: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dx" => vector.dx = string_to_int(&string_value),
                        b"dy" => vector.dy = string_to_int(&string_value),
                        b"dz" => vector.dz = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(vector)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.anchor?view=openxml-3.0.1
///
/// Example:
/// ```
/// <anchor x="123" y="23" z="10000"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAnchor {
    // Attributes	Description
    // x (X-Coordinate in 3D)	X-Coordinate in 3D space.
    pub x: Option<i64>,

    // y (Y-Coordinate in 3D)	Y-Coordinate in 3D space.
    pub y: Option<i64>,

    // z (Z-Coordinate in 3D)	Z-Coordinate in 3D space.
    pub z: Option<i64>,
}

impl XlsxAnchor {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut anchor = Self {
            x: None,
            y: None,
            z: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"x" => anchor.x = string_to_int(&string_value),
                        b"y" => anchor.y = string_to_int(&string_value),
                        b"z" => anchor.z = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(anchor)
    }
}
