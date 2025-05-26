use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::st_types::{STPositiveAngle, STPositivePercentage};
use crate::{helper::string_to_bool, raw::drawing::color::XlsxColorEnum};

use super::fill_rectangle::{XlsxFillToRectangle, XlsxTileRectangle};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:gradFill rotWithShape="1">
///     <a:gsLst>
///         <a:gs pos="0">
///             <a:schemeClr val="phClr">
///                 <a:satMod val="103000" />
///                 <a:lumMod val="102000" />
///                 <a:tint val="94000" />
///             </a:schemeClr>
///         </a:gs>
///         <a:gs pos="100000">
///             <a:schemeClr val="phClr">
///                 <a:lumMod val="99000" />
///                 <a:satMod val="120000" />
///                 <a:shade val="78000" />
///             </a:schemeClr>
///         </a:gs>
///     </a:gsLst>
///     <a:lin ang="5400000" scaled="0" />
/// </a:gradFill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGradientFill {
    // Child Elements
    /// The list of gradient stops that specifies the gradient colors and their relative positions in the color band.
    // tag: gsLst
    pub gs_lst: Option<Vec<XlsxGradientStop>>,

    /// specifies a linear gradient
    pub lin: Option<XlsxLinearGradientFill>,

    /// defines that a gradient fill follows a path
    // tag: path
    pub path: Option<XlsxPathGradientFill>,

    /// This element specifies a rectangular region of the shape to which the gradient is applied.
    // tag: tileRect
    pub tile_rect: Option<XlsxTileRectangle>,

    // attributes
    /// Specifies the direction(s) in which to flip the gradient while tiling.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tileflipvalues?view=openxml-3.0.1
    /// flip (Tile Flip)
    pub flip: Option<String>,

    /// Specifies if a fill rotates along with a shape when the shape is rotated.
    // tag: rotWithShape
    pub rotate_with_shape: Option<bool>,
}

impl XlsxGradientFill {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            path: None,
            gs_lst: None,
            lin: None,
            tile_rect: None,
            flip: None,
            rotate_with_shape: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"flip" => {
                            fill.flip = Some(string_value);
                        }
                        b"rotWithShape" => {
                            fill.rotate_with_shape = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gsLst" => {
                    fill.gs_lst = Some(load_gradient_stops(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lin" => {
                    fill.lin = Some(XlsxLinearGradientFill::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"path" => {
                    fill.path = Some(XlsxPathGradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tileRect" => {
                    fill.tile_rect = Some(XlsxTileRectangle::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gradFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(fill)
    }
}

pub(crate) fn load_gradient_stops(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<Vec<XlsxGradientStop>> {
    let mut gs_list: Vec<XlsxGradientStop> = vec![];

    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gs" => {
                let gs = XlsxGradientStop::load(reader, e)?;
                gs_list.push(gs);
            }

            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gsLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(gs_list)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientstop?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:gs pos="100000">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:gs>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGradientStop {
    // children
    pub color: Option<XlsxColorEnum>,

    // attribute
    /// Specifies where this gradient stop should appear in the color band.
    /// This position is specified in the range [0%, 100%], which corresponds to the beginning and the end of the color band respectively.
    pub pos: Option<STPositivePercentage>,
}

impl XlsxGradientStop {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut stop = Self {
            color: None,
            pos: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"pos" => {
                            stop.pos = string_to_unsignedint(&string_value);
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

        stop.color = XlsxColorEnum::load(reader, b"gs")?;

        Ok(stop)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineargradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:lin ang="5400000" scaled="0" />
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLinearGradientFill {
    // attributes
    /// Specifies the direction of color change for the gradient.
    pub ang: Option<STPositiveAngle>,

    /// Whether the gradient angle scales with the fill region.
    pub scaled: Option<bool>,
}

impl XlsxLinearGradientFill {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            ang: None,
            scaled: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"ang" => {
                            fill.ang = string_to_unsignedint(&string_value);
                        }
                        b"scaled" => {
                            fill.scaled = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(fill)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathgradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:path path="circle">
///     <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
/// </a:path>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPathGradientFill {
    // children
    /// defines the "focus" rectangle for the center shade, specified relative to the fill tile rectangle
    // tag: fillToRect
    pub fill_to_rect: Option<XlsxFillToRectangle>,

    // attributes
    /// Specifies the direction of color change for the gradient.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathshadevalues?view=openxml-3.0.1
    pub path: Option<String>,
}

impl XlsxPathGradientFill {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            fill_to_rect: None,
            path: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"path" => {
                            fill.path = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillToRect" => {
                    fill.fill_to_rect = Some(XlsxFillToRectangle::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"path" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(fill)
    }
}
