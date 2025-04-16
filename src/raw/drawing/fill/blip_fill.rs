use crate::excel::XmlReader;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    helper::{string_to_bool, string_to_int},
    raw::drawing::image::blip::Blip,
};

use super::rectangle::{FillRectangle, SourceRectangle};

/// Blip Fill (Picture Fill): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blipfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <p:blipFill>
///     <a:blip r:embed="rId2"/>
///     <a:stretch>
///         <a:fillRect b="10000" r="25000"/>
///     </a:stretch>
/// </p:blipFill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct BlipFill {
    //     Child Elements	Subclause
    // blip (Blip)	§20.1.8.13
    pub blip: Option<Blip>,
    // srcRect (Source Rectangle)	§20.1.8.55
    pub source_rect: Option<SourceRectangle>,
    // stretch (Stretch)	§20.1.8.56
    pub stretch: Option<Stretch>,
    // tile (Tile)
    pub tile: Option<Tile>,

    //  Attributes
    /// Specifies the DPI (dots per inch) used to calculate the size of the blip.
    /// If not present or zero, the DPI in the blip is used.
    // dpi (DPI Setting)
    pub dpi: Option<i64>,

    /// Specifies that the fill should rotate with the shape.
    // rotWithShape (Rotate With Shape)
    pub rot_with_shape: Option<bool>,
}

impl BlipFill {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            blip: None,
            source_rect: None,
            stretch: None,
            tile: None,
            // attributes
            dpi: None,
            rot_with_shape: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dpi" => {
                            fill.dpi = string_to_int(&string_value);
                        }
                        b"rotWithShape" => {
                            fill.rot_with_shape = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blip" => {
                    fill.blip = Some(Blip::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"srcRect" => {
                    fill.source_rect = Some(SourceRectangle::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"stretch" => {
                    fill.stretch = Some(Stretch::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tile" => {
                    fill.tile = Some(Tile::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"blipFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(fill)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.stretch?view=openxml-3.0.1
///
/// This element specifies a fill rectangle.
/// When stretching of an image is specified, a source rectangle, srcRect, is scaled to fit the specified fill rectangle.
// tag: stretch
#[derive(Debug, Clone, PartialEq)]
pub struct Stretch {
    // Child Elements	Subclause
    // fillRect (Fill Rectangle)
    pub fill_rectangle: Option<FillRectangle>,
}

impl Stretch {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut stretch = Self {
            fill_rectangle: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillRect" => {
                    stretch.fill_rectangle = Some(FillRectangle::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"stretch" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(stretch)
    }
}

// tag: tile
#[derive(Debug, Clone, PartialEq)]
pub struct Tile {
    // Attributes	Description
    /// Specifies where to align the first tile with respect to the shape.
    /// Alignment happens after the scaling, but before the additional offset.
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
    // algn (Alignment)
    pub alignment: Option<String>,

    /// Specifies the direction(s) in which to flip the source image while tiling.
    /// Images can be flipped horizontally, vertically, or in both directions to fill the entire region.
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tileflipvalues?view=openxml-3.0.1
    // flip (Tile Flipping)
    pub flip: Option<String>,

    /// Specifies the amount to horizontally scale the srcRect.
    // sx (Horizontal Ratio)
    pub sx: Option<i64>,

    /// Specifies the amount to vertically scale the srcRect.
    // sy (Vertical Ratio)
    pub sy: Option<i64>,

    /// Specifies additional horizontal offset after alignment.
    // tx (Horizontal Offset)
    pub tx: Option<i64>,

    /// Specifies additional vertical offset after alignment.
    // ty (Vertical Offset)
    pub ty: Option<i64>,
}

impl Tile {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut tile = Self {
            alignment: None,
            flip: None,
            sx: None,
            sy: None,
            tx: None,
            ty: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            tile.alignment = Some(string_value);
                        }
                        b"flip" => {
                            tile.flip = Some(string_value);
                        }
                        b"sx" => {
                            tile.sx = string_to_int(&string_value);
                        }
                        b"sy" => {
                            tile.sy = string_to_int(&string_value);
                        }
                        b"tx" => {
                            tile.tx = string_to_int(&string_value);
                        }
                        b"ty" => {
                            tile.ty = string_to_int(&string_value);
                        }

                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(tile)
    }
}
