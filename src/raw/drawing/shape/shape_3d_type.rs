use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::st_types::{STCoordinate, STPositiveCoordinate};
use crate::{helper::string_to_int, raw::drawing::color::XlsxColorEnum};

use super::bevel::{XlsxBevelBottom, XlsxBevelTop};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shape3dtype?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d extrusionH="165100" contourW="50800" prstMaterial="plastic">
///   <a:bevelT w="254000" h="254000"/>
///   <a:bevelB w="254000" h="254000"/>
///   <a:extrusionClr>
///     <a:srgbClr val="FF0000"/>
///   </a:extrusionClr>
///   <a:contourClr>
///     <a:schemeClr val="accent3"/>
///   </a:contourClr>
/// </a:sp3d>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShape3DType {
    // extLst (Extension List)	§20.1.2.2.15 No Supported

    // children
    // bevelB (Bottom Bevel)
    pub bevel_b: Option<XlsxBevelBottom>,

    // bevelT (Top Bevel)
    pub bevel_t: Option<XlsxBevelTop>,

    // contourClr (Contour Color)
    pub contour_clr: Option<XlsxColorEnum>,

    // extrusionClr (Extrusion Color)
    pub extrusion_clr: Option<XlsxColorEnum>,

    // attributes
    /// Extrusion Height
    // tag: extrusionH
    pub extrusion_h: Option<STPositiveCoordinate>,

    /// contour width
    // tag: contourW
    pub contour_w: Option<STPositiveCoordinate>,

    /// Preset Material Type
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetmaterialtypevalues?view=openxml-3.0.1
    // tag: prstMaterial
    pub prst_material: Option<String>,

    /// Shape Depth
    pub z: Option<STCoordinate>,
}

impl XlsxShape3DType {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut shape3d = Self {
            bevel_b: None,
            bevel_t: None,
            contour_clr: None,
            extrusion_clr: None,
            extrusion_h: None,
            contour_w: None,
            prst_material: None,
            z: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"extrusionH" => shape3d.extrusion_h = string_to_unsignedint(&string_value),
                        b"contourW" => shape3d.contour_w = string_to_unsignedint(&string_value),
                        b"prstMaterial" => shape3d.prst_material = Some(string_value),
                        b"z" => shape3d.z = string_to_int(&string_value),
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bevelB" => {
                    shape3d.bevel_b = Some(XlsxBevelBottom::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bevelT" => {
                    shape3d.bevel_t = Some(XlsxBevelTop::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"contourClr" => {
                    shape3d.contour_clr = XlsxColorEnum::load(reader, b"contourClr")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extrusionClr" => {
                    shape3d.extrusion_clr = XlsxColorEnum::load(reader, b"extrusionClr")?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sp3d" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(shape3d)
    }
}
