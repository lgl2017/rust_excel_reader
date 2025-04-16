use super::scheme::{
    color_scheme::ColorScheme, font_scheme::FontScheme, format_scheme::FormatScheme,
};
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.themeelements?view=openxml-3.0.1
///
/// This element contains the color scheme, font scheme, and format scheme elements which define the different formatting aspects of what a theme defines.
///
/// Example:
/// ```
/// <themeElements>
///   <clrScheme name="sample">
///   …  </clrScheme>
///     <fontScheme name="sample">
///   …  </fontScheme>
///     <fmtScheme name="sample">
///       <fillStyleLst>
///   …    </fillStyleLst>
///       <lnStyleLst>
///   …    </lnStyleLst>
///       <effectStyleLst>
///   …    </effectStyleLst>
///       <bgFillStyleLst>
///   …    </bgFillStyleLst>
///     </fmtScheme>
///   </themeElements>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct ThemeElements {
    // extLst (Extension List)	§20.1.2.2.15 Not Supported

    // children
    // tag: clrScheme
    pub color_scheme: Option<ColorScheme>,

    // tag: fmtScheme
    pub format_scheme: Option<FormatScheme>,

    // fontScheme (Font Scheme)
    pub font_scheme: Option<FontScheme>,
}

impl ThemeElements {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut scheme = Self {
            color_scheme: None,
            format_scheme: None,
            font_scheme: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrScheme" => {
                    let color_scheme = ColorScheme::load(reader, e)?;
                    scheme.color_scheme = Some(color_scheme);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fmtScheme" => {
                    let format_scheme = FormatScheme::load(reader, e)?;
                    scheme.format_scheme = Some(format_scheme);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fontScheme" => {
                    let font_scheme = FontScheme::load(reader, e)?;
                    scheme.font_scheme = Some(font_scheme);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"themeElements" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(scheme)
    }
}
