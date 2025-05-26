use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::st_types::{STAngle, STPercentage, STPositiveAngle, STPositiveCoordinate};
use crate::{
    helper::{string_to_bool, string_to_int},
    raw::drawing::color::XlsxColorEnum,
};

/// outerShdw (Outer Shadow Effect)
///
/// specifies an outer shadow effect.
///
///  Example:
/// ```
/// <a:outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0" >
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:outerShdw>
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outershadow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxOuterShadow {
    // children
    pub color: Option<XlsxColorEnum>,

    // attribute
    /// Specifies shadow alignment; alignment happens first, effectively setting the origin for scale, skew, and offset.
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
    pub algn: Option<String>,

    /// Specifies the blur radius
    //  tag: blurRad
    pub blur_rad: Option<STPositiveCoordinate>,

    /// Specifies the direction to offset the shadow as angle
    pub dir: Option<STPositiveAngle>,

    /// Specifies how far to offset the shadow
    pub dist: Option<STPositiveCoordinate>,

    /// Specifies the horizontal skew angle
    pub kx: Option<STAngle>,

    /// Specifies the vertical skew angle
    pub ky: Option<STAngle>,

    /// Specifies whether the shadow rotates with the shape if the shape is rotated
    // tag: rotWithShape
    pub rot_with_shape: Option<bool>,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip.
    pub sx: Option<STPercentage>,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub sy: Option<STPercentage>,
}

impl XlsxOuterShadow {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut shadow = Self {
            color: None,
            blur_rad: None,
            dir: None,
            dist: None,
            algn: None,
            kx: None,
            ky: None,
            rot_with_shape: None,
            sx: None,
            sy: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            shadow.algn = Some(string_value);
                        }
                        b"blurRad" => {
                            shadow.blur_rad = string_to_unsignedint(&string_value);
                        }
                        b"dir" => {
                            shadow.dir = string_to_unsignedint(&string_value);
                        }
                        b"dist" => {
                            shadow.dist = string_to_unsignedint(&string_value);
                        }
                        b"kx" => {
                            shadow.kx = string_to_int(&string_value);
                        }
                        b"ky" => {
                            shadow.ky = string_to_int(&string_value);
                        }
                        b"rotWithShape" => {
                            shadow.rot_with_shape = string_to_bool(&string_value);
                        }
                        b"sx" => {
                            shadow.sx = string_to_int(&string_value);
                        }
                        b"sy" => {
                            shadow.sy = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        shadow.color = XlsxColorEnum::load(reader, b"outerShdw")?;

        Ok(shadow)
    }
}
