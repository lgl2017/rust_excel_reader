use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::raw::drawing::st_types::{STAdjustAngle, STAdjustCoordinate};

use super::position::XlsxPosition;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlepolar?view=openxml-3.0.1
///
/// This element specifies a polar adjust handle for a custom shape.
///
/// Example
/// ```
/// <a:ahLst>
///     <a:ahPolar gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahPolar>
/// </a:ahLst>
/// ```
// tag: ahPolar
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAdjustHandlePolar {
    // Child Elements
    // pos (Shape Position Coordinate)
    pub position: Option<XlsxPosition>,

    // Attributes
    /// Specifies the name of the guide that is updated with the adjustment angle from this adjust handle.
    // gdRefAng (Angle Adjustment Guide)
    pub angle_guide_ref: Option<String>,

    /// Specifies the name of the guide that is updated with the adjustment radius from this adjust handle.
    // gdRefR (Radial Adjustment Guide)
    pub radial_guid_ref: Option<String>,

    /// Specifies the maximum angle position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move angularly.
    /// That is the maxAng and minAng are equal.
    // maxAng (Maximum Angle Adjustment)
    pub max_angle_adjustment: Option<STAdjustAngle>,

    /// Specifies the maximum radial position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move radially.
    /// That is the maxR and minR are equal.
    // maxR (Maximum Radial Adjustment)
    pub max_radial_adjustment: Option<STAdjustCoordinate>,

    /// Specifies the minimum angle position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move angularly.
    /// That is the maxAng and minAng are equal.
    // minAng (Minimum Angle Adjustment)
    pub min_angle_adjustment: Option<STAdjustAngle>,

    /// Specifies the minimum radial position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move radially.
    /// That is the maxR and minR are equal.
    // minR (Minimum Radial Adjustment)
    pub min_radial_adjustment: Option<STAdjustCoordinate>,
}

impl XlsxAdjustHandlePolar {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut polar = Self {
            position: None,
            angle_guide_ref: None,
            radial_guid_ref: None,
            max_angle_adjustment: None,
            max_radial_adjustment: None,
            min_angle_adjustment: None,
            min_radial_adjustment: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"gdRefAng" => polar.angle_guide_ref = Some(string_value),
                        b"gdRefR" => polar.radial_guid_ref = Some(string_value),
                        b"maxAng" => {
                            polar.max_angle_adjustment =
                                Some(STAdjustAngle::from_string(&string_value))
                        }
                        b"maxR" => {
                            polar.max_radial_adjustment =
                                Some(STAdjustCoordinate::from_string(&string_value))
                        }
                        b"minAng" => {
                            polar.min_angle_adjustment =
                                Some(STAdjustAngle::from_string(&string_value))
                        }
                        b"minR" => {
                            polar.min_radial_adjustment =
                                Some(STAdjustCoordinate::from_string(&string_value))
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pos" => {
                    polar.position = Some(XlsxPosition::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ahPolar" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(polar)
    }
}
