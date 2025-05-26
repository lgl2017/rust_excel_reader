use crate::{
    excel::XmlReader, helper::string_to_unsignedint, raw::drawing::st_types::STPositiveCoordinate,
};

use anyhow::bail;
use std::io::Read;

use arc_to::XlsxArcTo;
use close_shape_path::XlsxCloseShapePath;
use cubic_bezier_curve_to::XlsxCubicBezierCurveTo;
use line_to::XlsxLineTo;
use move_to::XlsxMoveTo;
use quad_bezier_curve_to::XlsxQuadraticBezierCurveTo;
use quick_xml::events::{BytesStart, Event};

use crate::helper::string_to_bool;

pub mod arc_to;
pub mod close_shape_path;
pub mod cubic_bezier_curve_to;
pub mod line_to;
pub mod move_to;
pub mod path_list;
pub mod path_point;
pub mod quad_bezier_curve_to;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.path?view=openxml-3.0.1
///
/// This element specifies a creation path consisting of a series of moves, lines and curves that when combined forms a geometric shape
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="3810000" h="3581400" fill="none" extrusionOk="0">
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
/// ```
// tag: path
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPath {
    // Child Elements
    pub paths: Option<Vec<XlsxPathTypeEnum>>,

    // Attributes
    /// Specifies that the use of 3D extrusions are possible on this path.
    /// This allows the generating application to know whether 3D extrusion can be applied in any form.
    /// If this attribute is omitted, then a value of 0, or false is assumed.
    // extrusionOk (3D Extrusion Allowed)
    pub extrusion_allowed: Option<bool>,

    /// Specifies how the corresponding path should be filled.
    /// If this attribute is omitted, a value of "norm" is assumed.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathfillmodevalues?view=openxml-3.0.1
    // fill (Path Fill)
    pub fill: Option<String>,

    /// Specifies the height, or maximum y coordinate that should be used for within the path coordinate system.
    /// This value determines the vertical placement of all points within the corresponding path as they are all calculated using this height attribute as the max y coordinate.
    // h (Path Height)
    pub height: Option<STPositiveCoordinate>,

    /// Specifies if the corresponding path should have a path stroke shown.
    /// This is a boolean value that affect the outline of the path.
    /// If this attribute is omitted, a value of true is assumed.
    // stroke (Path Stroke)
    pub stroke: Option<bool>,

    /// Specifies the width, or maximum x coordinate that should be used for within the path coordinate system.
    /// This value determines the horizontal placement of all points within the corresponding path as they are all calculated using this width attribute as the max x coordinate.
    // w (Path Width)
    pub width: Option<STPositiveCoordinate>,
}

impl XlsxPath {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut path = Self {
            paths: Some(XlsxPathTypeEnum::load_list(reader, b"path")?),
            extrusion_allowed: None,
            fill: None,
            height: None,
            stroke: None,
            width: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"extrusionOk" => path.extrusion_allowed = string_to_bool(&string_value),
                        b"fill" => path.fill = Some(string_value),
                        b"h" => path.height = string_to_unsignedint(&string_value),
                        b"stroke" => path.stroke = string_to_bool(&string_value),
                        b"w" => path.width = string_to_unsignedint(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(path)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxPathTypeEnum {
    Arc(XlsxArcTo),
    Close(XlsxCloseShapePath),
    CubicBezier(XlsxCubicBezierCurveTo),
    Line(XlsxLineTo),
    Move(XlsxMoveTo),
    QuadBezier(XlsxQuadraticBezierCurveTo),
}

impl XlsxPathTypeEnum {
    pub(crate) fn load_list(
        reader: &mut XmlReader<impl Read>,
        tag: &[u8],
    ) -> anyhow::Result<Vec<Self>> {
        let mut paths: Vec<XlsxPathTypeEnum> = vec![];
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    if let Some(fill) = XlsxPathTypeEnum::load_helper(reader, e)? {
                        paths.push(fill);
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(paths)
    }

    fn load_helper(
        reader: &mut XmlReader<impl Read>,
        e: &BytesStart,
    ) -> anyhow::Result<Option<Self>> {
        match e.local_name().as_ref() {
            b"arcTo" => {
                return Ok(Some(XlsxPathTypeEnum::Arc(XlsxArcTo::load(e)?)));
            }
            b"close" => {
                return Ok(Some(XlsxPathTypeEnum::Close(true)));
            }
            b"cubicBezTo" => {
                return Ok(Some(XlsxPathTypeEnum::CubicBezier(
                    XlsxCubicBezierCurveTo::load(reader)?,
                )));
            }
            b"lnTo" => {
                return Ok(Some(XlsxPathTypeEnum::Line(XlsxLineTo::load(reader)?)));
            }
            b"moveTo" => {
                return Ok(Some(XlsxPathTypeEnum::Move(XlsxMoveTo::load(reader)?)));
            }
            b"quadBezTo" => {
                return Ok(Some(XlsxPathTypeEnum::QuadBezier(
                    XlsxQuadraticBezierCurveTo::load(reader)?,
                )));
            }
            _ => return Ok(None),
        }
    }
}
