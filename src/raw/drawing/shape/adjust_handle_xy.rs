use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::common_types::AdjustCoordinate;

use super::position::Position;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlexy?view=openxml-3.0.1
///
/// This element specifies an XY-based adjust handle for a custom shape.
///
/// Example
/// ```
/// <a:ahLst>
///     <a:ahXY gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahXY>
/// </a:ahLst>
/// ```
// tag: ahPolar
#[derive(Debug, Clone, PartialEq)]
pub struct AdjustHandleXY {
    // Child Elements
    // pos (Shape Position Coordinate)
    pub position: Option<Position>,

    // Attributes
    /// Specifies the name of the guide that is updated with the adjustment x position from this adjust handle.
    // gdRefX (Horizontal Adjustment Guide)
    pub horizontal_guide_ref: Option<String>,

    /// Specifies the name of the guide that is updated with the adjustment y position from this adjust handle.
    // gdRefY (Vertical Adjustment Guide)
    pub vertical_guide_ref: Option<String>,

    /// Specifies the maximum horizontal position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the x direction.
    /// That is the maxX and minX are equal.
    // maxX (Maximum Horizontal Adjustment)
    pub max_horizaontal_adjustment: Option<AdjustCoordinate>,

    /// Specifies the maximum vertical position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the y direction.
    /// That is the maxY and minY are equal.
    // maxY (Maximum Vertical Adjustment)
    pub max_vertical_adjustment: Option<AdjustCoordinate>,

    /// Specifies the minimum horizontal position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the x direction.
    /// That is the maxX and minX are equal.
    // minX (Minimum Horizontal Adjustment)
    pub min_horizaontal_adjustment: Option<AdjustCoordinate>,

    /// Specifies the minimum vertical position that is allowed for this adjustment handle.
    /// If this attribute is omitted, then it is assumed that this adjust handle cannot move in the y direction.
    /// That is the maxY and minY are equal.
    // minY (Minimum Vertical Adjustment)
    pub min_vertical_adjustment: Option<AdjustCoordinate>,
}

impl AdjustHandleXY {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut polar = Self {
            position: None,
            horizontal_guide_ref: None,
            vertical_guide_ref: None,
            max_horizaontal_adjustment: None,
            max_vertical_adjustment: None,
            min_horizaontal_adjustment: None,
            min_vertical_adjustment: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"gdRefX" => polar.horizontal_guide_ref = Some(string_value),
                        b"gdRefY" => polar.vertical_guide_ref = Some(string_value),
                        b"maxX" => {
                            polar.max_horizaontal_adjustment =
                                Some(AdjustCoordinate::from_string(&string_value))
                        }
                        b"maxY" => {
                            polar.max_vertical_adjustment =
                                Some(AdjustCoordinate::from_string(&string_value))
                        }
                        b"minX" => {
                            polar.min_horizaontal_adjustment =
                                Some(AdjustCoordinate::from_string(&string_value))
                        }
                        b"minY" => {
                            polar.min_vertical_adjustment =
                                Some(AdjustCoordinate::from_string(&string_value))
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
                    polar.position = Some(Position::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ahXY" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(polar)
    }
}
