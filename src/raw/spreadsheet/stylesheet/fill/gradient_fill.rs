use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader, helper::string_to_float, raw::spreadsheet::stylesheet::color::Color,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <fill>
///     <gradientFill degree="90">
///         <stop position="0">
///             <color rgb="FF92D050"/>
///         </stop>
///         <stop position="1">
///             <color rgb="FF0070C0"/>
///         </stop>
///     </gradientFill>
/// </fill>
/// <fill>
///     <gradientFill type="path" left="0.2" right="0.8" top="0.2" bottom="0.8">
///         <stop position="0">
///             <color theme="0"/>
///         </stop>
///         <stop position="1">
///             <color theme="4"/>
///         </stop>
///     </gradientFill>
/// </fill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct GradientFill {
    // attributes
    /// values ranging from 0 to 1.
    /// Specifies the position of the bottom edge of the inner rectangle
    /// For bottom, 0 means the bottom edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
    pub bottom: Option<f64>,

    /// values ranging from 0 to 1.
    /// Specifies the position of the left edge of the inner rectangle
    /// For left, 0 means the left edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
    pub left: Option<f64>,

    /// values ranging from 0 to 1.
    /// Specifies the position of the right edge of the inner rectangle
    /// For right, 0 means the right edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
    pub right: Option<f64>,

    /// values ranging from 0 to 1.
    /// Specifies the position of the top edge of the inner rectangle
    /// For top, 0 means the top edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
    pub top: Option<f64>,

    /// Angle of the linear gradient
    pub degree: Option<f64>,

    /// Type of gradient fill.
    /// Allowed value: linear, path
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientvalues?view=openxml-3.0.1
    pub r#type: Option<String>,

    // children
    pub stop: Option<Vec<GradientStop>>,
}

impl GradientFill {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            bottom: None,
            left: None,
            right: None,
            top: None,
            degree: None,
            r#type: None,
            stop: None,
        };
        let mut stops: Vec<GradientStop> = vec![];

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"bottom" => {
                            fill.bottom = string_to_float(&string_value);
                        }
                        b"left" => {
                            fill.left = string_to_float(&string_value);
                        }
                        b"right" => {
                            fill.right = string_to_float(&string_value);
                        }
                        b"top" => {
                            fill.top = string_to_float(&string_value);
                        }
                        b"degree" => {
                            fill.degree = string_to_float(&string_value);
                        }
                        b"type" => {
                            fill.r#type = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"stop" => {
                    let stop = GradientStop::load(reader, e)?;
                    stops.push(stop);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gradientFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        fill.stop = Some(stops);

        Ok(fill)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.gradientstop?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct GradientStop {
    // attributes
    /// Position information for this gradient stop
    ///
    /// Interpreted exactly like gradientFill left, right, bottom, top.
    /// The position indicated here indicates the point where the color is pure.
    /// Before and and after this position the color can be in transition (or pure, depending on if this is the last stop or not).
    pub position: Option<f64>,

    // children
    pub color: Option<Color>,
}

impl GradientStop {
    pub fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut stop = Self {
            position: None,
            color: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"position" => {
                            stop.position = string_to_float(&string_value);
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
                    let color: Color = Color::load(e)?;
                    stop.color = Some(color);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"stop" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(stop)
    }
}
