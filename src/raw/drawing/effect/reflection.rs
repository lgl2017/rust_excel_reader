use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.reflection?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct Reflection {
    /// attributes
    /// Specifies shadow alignment; alignment happens first, effectively setting the origin for scale, skew, and offset.
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectanglealignmentvalues?view=openxml-3.0.1
    pub algn: Option<String>,

    /// Specifies the blur radius
    //  tag: blurRad
    pub blur_rad: Option<i64>,

    /// Specifies the direction to offset the shadow as angle
    pub dir: Option<i64>,

    /// Specifies how far to offset the shadow
    pub dist: Option<i64>,

    /// Specifies the ending reflection opacity.
    // endA (End Alpha)
    pub end_a: Option<i64>,

    /// Specifies the end position (along the alpha gradient ramp) of the end alpha value.
    // endPos (End Position)
    pub end_pos: Option<i64>,

    ///Specifies the direction to offset the reflection.
    // fadeDir (Fade Direction)
    pub fade_dir: Option<i64>,

    /// Specifies the horizontal skew angle
    pub kx: Option<i64>,

    /// Specifies the vertical skew angle
    pub ky: Option<i64>,

    /// Specifies whether the shadow rotates with the shape if the shape is rotated
    // tag: rotWithShape
    pub rot_with_shape: Option<bool>,

    /// starting reflection opacity.
    // stA (Start Opacity)
    pub st_a: Option<i64>,

    /// Specifies the start position (along the alpha gradient ramp) of the start alpha value.
    // stPos (Start Position)
    pub st_pos: Option<i64>,

    /// Specifies the horizontal scaling factor; negative scaling causes a flip.
    pub sx: Option<i64>,

    /// Specifies the vertical scaling factor; negative scaling causes a flip.
    pub sy: Option<i64>,
}

impl Reflection {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reflection = Self {
            algn: None,
            blur_rad: None,
            dir: None,
            dist: None,
            end_a: None,
            end_pos: None,
            fade_dir: None,
            kx: None,
            ky: None,
            rot_with_shape: None,
            st_a: None,
            st_pos: None,
            sx: None,
            sy: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            reflection.algn = Some(string_value);
                        }
                        b"blurRad" => {
                            reflection.blur_rad = string_to_int(&string_value);
                        }
                        b"dir" => {
                            reflection.dir = string_to_int(&string_value);
                        }
                        b"dist" => {
                            reflection.dist = string_to_int(&string_value);
                        }
                        b"endA" => {
                            reflection.end_a = string_to_int(&string_value);
                        }
                        b"endPos" => {
                            reflection.end_pos = string_to_int(&string_value);
                        }
                        b"fadeDir" => {
                            reflection.fade_dir = string_to_int(&string_value);
                        }
                        b"kx" => {
                            reflection.kx = string_to_int(&string_value);
                        }
                        b"ky" => {
                            reflection.ky = string_to_int(&string_value);
                        }
                        b"rotWithShape" => {
                            reflection.rot_with_shape = string_to_bool(&string_value);
                        }
                        b"stA" => {
                            reflection.st_a = string_to_int(&string_value);
                        }
                        b"stPos" => {
                            reflection.st_pos = string_to_int(&string_value);
                        }
                        b"sx" => {
                            reflection.sx = string_to_int(&string_value);
                        }
                        b"sy" => {
                            reflection.sy = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(reflection)
    }
}
