use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::raw::drawing::st_types::STAdjustCoordinate;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectangle?view=openxml-3.0.1
///
/// This element specifies the rectangular bounding box for text within a `custGeom` shape.
/// The default for this rectangle is the bounding box for the shape.
/// This can be modified using this elements four attributes to inset or extend the text bounding box.
///
/// rect (Shape Text Rectangle)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShapeTextRectangle {
    // attributes
    /// b (Bottom Position)
    ///
    /// Specifies the y coordinate of the bottom edge for a shape text rectangle.
    /// The units for this edge is specified in EMUs as the positioning here is based on the shape coordinate system.
    pub b: Option<STAdjustCoordinate>,

    /// l (Left)
    ///
    /// Specifies the x coordinate of the left edge for a shape text rectangle.
    /// The units for this edge is specified in EMUs as the positioning here is based on the shape coordinate system.
    pub l: Option<STAdjustCoordinate>,

    /// r (Right)
    ///
    /// Specifies the x coordinate of the right edge for a shape text rectangle.
    /// The units for this edge is specified in EMUs as the positioning here is based on the shape coordinate system.
    pub r: Option<STAdjustCoordinate>,

    /// t (Top)
    ///
    /// Specifies the y coordinate of the top edge for a shape text rectangle.
    /// The units for this edge is specified in EMUs as the positioning here is based on the shape coordinate system.
    pub t: Option<STAdjustCoordinate>,
}

impl XlsxShapeTextRectangle {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut rect = Self {
            b: None,
            l: None,
            r: None,
            t: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"b" => {
                            rect.b = Some(STAdjustCoordinate::from_string(&string_value));
                        }
                        b"l" => {
                            rect.l = Some(STAdjustCoordinate::from_string(&string_value));
                        }
                        b"r" => {
                            rect.r = Some(STAdjustCoordinate::from_string(&string_value));
                        }
                        b"t" => {
                            rect.t = Some(STAdjustCoordinate::from_string(&string_value));
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(rect)
    }
}
