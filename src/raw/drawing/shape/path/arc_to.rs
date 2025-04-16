use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::common_types::{AdjustAngle, AdjustCoordinate};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.arcto?view=openxml-3.0.1
///
/// This element specifies the existence of an arc within a shape path.
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:arcTo hR="123"/>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct ArcTo {
    // Attributes
    /// This attribute specifies the height radius of the supposed circle being used to draw the arc. This gives the circle a total height of (2 * hR). This total height could also be called it's vertical diameter as it is the diameter for the y axis only.
    // hR (Shape Arc Height Radius)
    pub height_radius: Option<AdjustCoordinate>,

    /// Specifies the start angle for an arc. This angle specifies what angle along the supposed circle path is used as the start position for drawing the arc. This start angle is locked to the last known pen position in the shape path. Thus guaranteeing a continuos shape path.
    // stAng (Shape Arc Start Angle)
    pub start_angle: Option<AdjustAngle>,

    /// Specifies the swing angle for an arc. This angle specifies how far angle-wise along the supposed cicle path the arc is extended. The extension from the start angle is always in the clockwise direction around the supposed circle.
    // swAng (Shape Arc Swing Angle)
    pub swing_angle: Option<AdjustAngle>,

    /// This attribute specifies the width radius of the supposed circle being used to draw the arc. This gives the circle a total width of (2 * wR). This total width could also be called it's horizontal diameter as it is the diameter for the x axis only.
    // wR (Shape Arc Width Radius)
    pub width_radius: Option<AdjustCoordinate>,
}

impl ArcTo {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut arc = Self {
            height_radius: None,
            start_angle: None,
            swing_angle: None,
            width_radius: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"hR" => {
                            arc.height_radius = Some(AdjustCoordinate::from_string(&string_value))
                        }
                        b"stAng" => arc.start_angle = Some(AdjustAngle::from_string(&string_value)),
                        b"swAng" => arc.swing_angle = Some(AdjustAngle::from_string(&string_value)),
                        b"wR" => {
                            arc.width_radius = Some(AdjustCoordinate::from_string(&string_value))
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(arc)
    }
}
