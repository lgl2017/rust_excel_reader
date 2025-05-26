use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader, helper::string_to_bool,
    raw::drawing::non_visual_properties::non_visual_connection_shape_properties::XlsxNonVisualConnectionShapeProperties,
};

use super::{shape_properties::XlsxShapeProperties, shape_style::XlsxShapeStyle};

/// - ConnectionShape: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionshape?view=openxml-3.0.1
/// - SpreadSheet.ConnectionShape: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.connectionshape?view=openxml-3.0.1
///
/// This element specifies a connection shape that is used to connect two sp elements.
/// Once a connection is specified using a cxnSp, it is left to the generating application to determine the exact path the connector takes.
/// That is the connector routing algorithm is left up to the generating application as the desired path might be different depending on the specific needs of the application.
///
/// Example:
/// ```
/// <xdr:cxnSp macro="">
///     <xdr:nvCxnSpPr>
///         <xdr:cNvPr id="5" name="Straight Arrow Connector 4">
///         </xdr:cNvPr>
///         <xdr:cNvCxnSpPr>
///             <a:stCxn id="3" idx="1" />
///             <a:endCxn id="2" idx="4" />
///         </xdr:cNvCxnSpPr>
///     </xdr:nvCxnSpPr>
///     <xdr:spPr>
///         <a:xfrm flipH="1" flipV="1">
///             <a:off x="2311398" y="1505943" />
///             <a:ext cx="1333502" cy="741957" />
///         </a:xfrm>
///         <a:prstGeom prst="straightConnector1">
///             <a:avLst />
///         </a:prstGeom>
///         <a:ln>
///             <a:solidFill>
///                 <a:schemeClr val="accent1" />
///             </a:solidFill>
///             <a:bevel />
///             <a:tailEnd type="triangle" w="med" len="lg" />
///         </a:ln>
///     </xdr:spPr>
///     <xdr:style>
///     … </xdr:style>
/// </xdr:cxnSp>
/// ```
///
/// cxnSp (Connection Shape)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxConnectionShape {
    // Child Elements	Subclause
    // extLst (Extension List)	Not supported

    // nvCxnSpPr (Non-Visual Properties for a Connection Shape)
    pub non_visual_connection_shape_properties: Option<XlsxNonVisualConnectionShapeProperties>,

    // spPr (Shape Properties)
    pub shape_properties: Option<XlsxShapeProperties>,

    // style (Shape Style)
    pub shape_style: Option<XlsxShapeStyle>,

    // attributes for SpreadSheet.ConnectionShape
    /// Macro
    ///
    /// Reference to Custom Function
    pub r#macro: Option<String>,

    /// Published
    ///
    /// Publish to Server Flag
    pub published: Option<bool>,
}

impl XlsxConnectionShape {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut connection_shape = Self {
            non_visual_connection_shape_properties: None,
            shape_properties: None,
            shape_style: None,
            r#macro: None,
            published: None,
        };
        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"macro" => {
                            connection_shape.r#macro = Some(string_value);
                            break;
                        }
                        b"fPublished" => {
                            connection_shape.published = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvCxnSpPr" => {
                    connection_shape.non_visual_connection_shape_properties =
                        Some(XlsxNonVisualConnectionShapeProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                    connection_shape.shape_properties = Some(XlsxShapeProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"style" => {
                    connection_shape.shape_style = Some(XlsxShapeStyle::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cxnSp" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at XlsxConnectionShape: `cxnSp`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(connection_shape)
    }
}
