use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{excel::XmlReader, helper::string_to_bool};

use super::color::XlsxColor;

pub type XlsxBorders = Vec<XlsxBorder>;

/// Borders: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borders?view=openxml-3.0.1
pub(crate) fn load_borders(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxBorders> {
    let mut buf: Vec<u8> = Vec::new();
    let mut borders: Vec<XlsxBorder> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                let border = XlsxBorder::load(reader, e)?;
                borders.push(border);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"borders" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(borders)
}

/// Border: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.border?view=openxml-3.0.1
///
/// Expresses a single set of cell border formats (left, right, top, bottom, diagonal).
/// Color is optional. When missing, 'automatic' is implied.
///
/// Example
/// ```
/// <borders count="3">
///     <border>
///         <left/>
///         <right/>
///         <top/>
///         <bottom/>
///         <diagonal/>
///     </border>
///     <border>
///         <left/>
///         <right style="medium">
///             <color indexed="64"/>
///         </right>
///         <top/>
///         <bottom style="thin">
///             <color indexed="64"/>
///         </bottom>
///         <diagonal/>
///     </border>
///     <border>
///         <left/>
///         <right/>
///         <top style="double">
///             <color auto="1"/>
///         </top>
///         <bottom/>
///         <diagonal/>
///     </border>
/// </borders>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBorder {
    // attributes
    // xml tag: diagonalDown
    /// A boolean value indicating if the cell's diagonal border includes a diagonal line,
    /// starting at the top left corner of the cell and moving down to the bottom right corner of the cell.
    pub diagonal_down: Option<bool>,

    // xml tag: diagonalUp
    /// A boolean value indicating if the cell's diagonal border includes a diagonal line,
    /// starting at the bottom left corner of the cell and moving up to the top right corner of the cell.
    pub diagonal_up: Option<bool>,

    /// A boolean value indicating if left, right, top, and bottom borders should be applied only to outside borders of a cell range.
    pub outline: Option<bool>,

    // children
    // xml tag: left or start
    pub left: Option<XlsxLeftBorder>,
    // xml tag: right or end
    pub right: Option<XlsxRightBorder>,
    pub top: Option<XlsxTopBorder>,
    pub bottom: Option<XlsxBottomBorder>,
    pub diagonal: Option<XlsxDiagonalBorder>,
}

impl XlsxBorder {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut border = Self {
            diagonal_down: None,
            diagonal_up: None,
            outline: None,
            left: None,
            right: None,
            top: None,
            bottom: None,
            diagonal: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"diagonalDown" => {
                            border.diagonal_down = string_to_bool(&string_value);
                        }
                        b"diagonalUp" => {
                            border.diagonal_down = string_to_bool(&string_value);
                        }
                        b"outline" => {
                            border.diagonal_down = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"left" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"left")?;
                    border.left = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"start" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"start")?;
                    border.left = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"right" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"right")?;
                    border.right = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"end" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"end")?;
                    border.right = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"top" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"top")?;
                    border.top = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bottom" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"bottom")?;
                    border.bottom = Some(border_style);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"diagonal" => {
                    let border_style = XlsxBorderStyle::load(reader, e, b"diagonal")?;
                    border.diagonal = Some(border_style);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"border" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(border)
    }
}

/// RightBorder (Semantically equivalent to end): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.rightborder?view=openxml-3.0.1
pub type XlsxRightBorder = XlsxBorderStyle;

/// LeftBorder (Semantically equivalent to start): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.leftborder?view=openxml-3.0.1
pub type XlsxLeftBorder = XlsxBorderStyle;

/// TopBorder: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.topborder?view=openxml-3.0.1
pub type XlsxTopBorder = XlsxBorderStyle;

/// BottomBorder: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bottomborder?view=openxml-3.0.1
pub type XlsxBottomBorder = XlsxBorderStyle;

/// DiagonalBorder: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.diagonalborder?view=openxml-3.0.1
pub type XlsxDiagonalBorder = XlsxBorderStyle;

/// EndBorder(Trailing Edge Border): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.endborder?view=openxml-3.0.1
pub type XlsxEndBorder = XlsxBorderStyle;

/// StartBorder(Leading Edge Border): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.startborder?view=openxml-3.0.1
pub type XlsxStartBorder = XlsxBorderStyle;

/// HorizontalBorder(Horizontal inner Border): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.horizontalborder?view=openxml-3.0.1
pub type XlsxHorizontalBorder = XlsxBorderStyle;

/// VerticalBorder(Vertical Inner Border): https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticalborder?view=openxml-3.0.1
pub type XlsxVerticalBorder = XlsxBorderStyle;

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBorderStyle {
    // attributes
    /// The line style for this border
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borderstylevalues?view=openxml-3.0.1
    pub style: Option<String>,

    // children
    /// Data Bar Color
    pub color: Option<XlsxColor>,
}

impl XlsxBorderStyle {
    pub fn load(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
        tag: &[u8],
    ) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut border_style = Self {
            color: None,
            style: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"style" => {
                            border_style.style = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                    let color = XlsxColor::load(e)?;
                    border_style.color = Some(color);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(border_style)
    }
}
