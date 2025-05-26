use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{
    helper::string_to_int,
    raw::drawing::st_types::{STAngle, STCoordinate, STPercentage},
};

/// xfrm (Transform Effect)
///
/// This element specifies a transform effect. The transform is applied to each point in the shape's geometry using the following matrix:
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transformeffect?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTransformEffect {
    /// Specifies the horizontal skew angle
    pub kx: Option<STAngle>,

    /// Specifies the vertical skew angle
    pub ky: Option<STAngle>,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip.
    pub sx: Option<STPercentage>,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub sy: Option<STPercentage>,

    /// Specifies an amount by which to shift the object along the x-axis
    pub tx: Option<STCoordinate>,

    /// Specifies an amount by which to shift the object along the y-axis
    pub ty: Option<STCoordinate>,
}

impl XlsxTransformEffect {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut transform = Self {
            kx: None,
            ky: None,
            sx: None,
            sy: None,
            tx: None,
            ty: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"kx" => {
                            transform.kx = string_to_int(&string_value);
                        }
                        b"ky" => {
                            transform.ky = string_to_int(&string_value);
                        }
                        b"sx" => {
                            transform.sx = string_to_int(&string_value);
                        }
                        b"sy" => {
                            transform.sy = string_to_int(&string_value);
                        }
                        b"tx" => {
                            transform.tx = string_to_int(&string_value);
                        }
                        b"ty" => {
                            transform.ty = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(transform)
    }
}
