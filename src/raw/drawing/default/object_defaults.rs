use super::{
    line_default::{load_line_default, XlsxLineDefault},
    shape_default::{load_shape_default, XlsxShapeDefault},
    text_default::{load_text_default, XlsxTextDefault},
};
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.objectdefaults?view=openxml-3.0.1
/// This element allows for the definition of default shape, line, and textbox formatting properties.
/// An application can use this information to format a shape (or text) initially on insertion into a document.
///
/// Example:
/// ```
/// <a:objectDefaults>
///     <a:spDef>
///         <a:spPr>
///             <a:ln>
///                 <a:noFill />
///             </a:ln>
///         </a:spPr>
///         <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr" anchorCtr="0" />
///         <a:lstStyle>
///             <a:defPPr algn="ctr">
///                 <a:defRPr sz="1600">
///                     <a:solidFill>
///                         <a:schemeClr val="bg1" />
///                     </a:solidFill>
///                 </a:defRPr>
///             </a:defPPr>
///         </a:lstStyle>
///         <a:style>
///             <a:lnRef idx="2">
///                 <a:schemeClr val="accent1">
///                     <a:shade val="50000" />
///                 </a:schemeClr>
///             </a:lnRef>
///             <a:fillRef idx="1">
///                 <a:schemeClr val="accent1" />
///             </a:fillRef>
///             <a:effectRef idx="0">
///                 <a:schemeClr val="accent1" />
///             </a:effectRef>
///             <a:fontRef idx="minor">
///                 <a:schemeClr val="lt1" />
///             </a:fontRef>
///         </a:style>
///     </a:spDef>
///     <lnDef>
///         <spPr/>
///         <bodyPr/>
///         <lstStyle/>
///         <style>
///             <lnRef idx="1">
///                 <schemeClr val="accent2"/>
///             </lnRef>
///             <fillRef idx="0">
///             <schemeClr val="accent2"/>
///             </fillRef>
///             <effectRef idx="0">
///                 <schemeClr val="accent2"/>
///             </effectRef>
///             <fontRef idx="minor">
///                 <schemeClr val="tx1"/>
///             </fontRef>
///         </style>
///     </lnDef>
/// </a:objectDefaults>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxObjectDefaults {
    // extLst (Extension List)	§20.1.2.2.15 Not supporte

    // Child Elements
    // lnDef (Line Default)	§20.1.4.1.20
    pub line_defaults: Option<Box<XlsxLineDefault>>,

    // spDef (Shape Default)	§20.1.4.1.27
    pub shape_defaults: Option<Box<XlsxShapeDefault>>,

    // txDef (Text Default)
    pub text_defaults: Option<Box<XlsxTextDefault>>,
}

impl XlsxObjectDefaults {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut defaults = Self {
            line_defaults: None,
            shape_defaults: None,
            text_defaults: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnDef" => {
                    defaults.line_defaults = Some(Box::new(load_line_default(reader)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spDef" => {
                    defaults.shape_defaults = Some(Box::new(load_shape_default(reader)?));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txDef" => {
                    defaults.text_defaults = Some(Box::new(load_text_default(reader)?));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"objectDefaults" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(defaults)
    }
}
