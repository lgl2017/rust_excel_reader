use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::{
    adjust_handle_list::AdjustHandleList,
    adjust_value_list::{load_adjust_value_list, AdjustValueList},
    connection_site_list::{load_connection_site_list, ConnectionSiteList},
    path::path_list::{load_path_list, PathList},
    rectangle::Rectangle,
    shape_guide_list::{load_shape_guide_list, ShapeGuideList},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customgeometry?view=openxml-3.0.1
///
/// This element specifies the existence of a custom geometric shape.
///
/// Example
/// ```
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst/>
///   <a:ahLst/>
///   <a:cxnLst>
///     <a:cxn ang="0">
///         <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///         <a:pos x="1705233" y="679622"/>
///     </a:cxn>
///   </a:cxnLst>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:moveTo>
///         <a:pt x="0" y="1261641"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2650602" y="1261641"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1226916" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
// tag: custGeom
#[derive(Debug, Clone, PartialEq)]
pub struct CustomGeometry {
    // Child Elements
    // ahLst (List of Shape Adjust Handles)	§20.1.9.1
    pub adjust_handle_list: Option<AdjustHandleList>,

    // avLst (List of Shape Adjust Values)	§20.1.9.5
    pub adjust_value_list: Option<AdjustValueList>,

    // cxnLst (List of Shape Connection Sites)	§20.1.9.10
    pub connection_site_list: Option<ConnectionSiteList>,

    // gdLst (List of Shape Guides)	§20.1.9.12
    pub shape_guide_list: Option<ShapeGuideList>,

    // pathLst (List of Shape Paths)	§20.1.9.16
    pub path_list: Option<PathList>,

    // rect (Shape Text Rectangle)
    pub text_rectangle: Option<Rectangle>,
}

impl CustomGeometry {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut geom = Self {
            adjust_handle_list: None,
            adjust_value_list: None,
            connection_site_list: None,
            shape_guide_list: None,
            path_list: None,
            text_rectangle: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ahLst" => {
                    geom.adjust_handle_list = Some(AdjustHandleList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"avLst" => {
                    geom.adjust_value_list = Some(load_adjust_value_list(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cxnLst" => {
                    geom.connection_site_list = Some(load_connection_site_list(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gdLst" => {
                    geom.shape_guide_list = Some(load_shape_guide_list(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pathLst" => {
                    geom.path_list = Some(load_path_list(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rect" => {
                    geom.text_rectangle = Some(Rectangle::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"custGeom" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(geom)
    }
}
