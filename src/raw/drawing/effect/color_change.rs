use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{helper::string_to_bool, raw::drawing::color::XlsxColorEnum};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorchange?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:clrChange useA="1">
///     <a:clrFrom>
///         <a:schemeClr val="phClr" />
///     </a:clrFrom>
///     <a:clrTo>
///         <a:schemeClr val="phClr" />
///     </a:clrTo>
/// </a:clrChange>
/// ```
// tag: clrChange
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxColorChange {
    // Children
    // tag: clrFrom (Change Color From)	§20.1.8.17
    pub color_from: Option<XlsxColorFrom>,
    // tag: clrTo (Change Color To)
    pub color_to: Option<XlsxColorTo>,

    // attributes
    /// Specifies whether alpha values are considered for the effect.
    /// Effect alpha values are considered if useA is true, else they are ignored.
    // tag: useA (Consider Alpha Values)
    pub use_a: Option<bool>,
}

impl XlsxColorChange {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut change = Self {
            color_from: None,
            color_to: None,
            use_a: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"useA" => {
                            change.use_a = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrTo" => {
                    change.color_from = XlsxColorFrom::load(reader, b"clrFrom")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrTo" => {
                    change.color_to = XlsxColorTo::load(reader, b"clrTo")?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"clrChange" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(change)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorfrom?view=openxml-3.0.1
/// specifies a color getting removed in a color change effect.
/// It is the "from" or source input color.
///
/// Example:
/// ```
/// <a:clrFrom>
///     <a:schemeClr val="phClr" />
/// </a:clrFrom>
/// ```
// tag: clrFrom
pub type XlsxColorFrom = XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorto?view=openxml-3.0.1
/// specifies the color which replaces the clrFrom in a clrChange effect.
/// This is the "target" or "to" color in the color change effect///
/// Example:
/// ```
/// <a:clrTo>
///     <a:schemeClr val="phClr" />
/// </a:clrTo>
/// ```
// tag: clrTo
pub type XlsxColorTo = XlsxColorEnum;
