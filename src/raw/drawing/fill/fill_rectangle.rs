use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STPercentage};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tilerectangle?view=openxml-3.0.1
pub type XlsxTileRectangle = XlsxFillRectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.sourcerectangle?view=openxml-3.0.1
///
/// This element specifies the portion of the blip used for the fill.
pub type XlsxSourceRectangle = XlsxFillRectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filltorectangle?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
/// ```
pub type XlsxFillToRectangle = XlsxFillRectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillrectangle?view=openxml-3.0.1
///
/// This element specifies a fill rectangle.
/// When stretching of an image is specified, a source rectangle, srcRect, is scaled to fit the specified fill rectangle.
///
/// Example:
/// ```
/// <a:blipFill>
///   <a:blip r:embed="rId2"/>
///   <a:stretch>
///       <a:fillRect b="10000" r="25000"/>
///   </a:stretch>
/// </a:blipFill>
/// ```
/// The above image is stretched to fill the entire rectangle except for the bottom 10% and rightmost 25%.
///
/// fillRect (Fill Rectangle)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFillRectangle {
    // attributes
    /// Specifies the bottom edge of the rectangle in percentage.
    ///
    /// b (Bottom Offset)
    pub b: Option<STPercentage>,

    /// Specifies the left edge of the rectangle
    ///
    /// l (Left Offset)
    pub l: Option<STPercentage>,

    /// Specifies the right edge of the rectangle
    ///
    /// r (Right Offset)
    pub r: Option<STPercentage>,

    /// Specifies the top edge of the rectangle
    ///
    /// t (Top Offset)
    pub t: Option<STPercentage>,
}

impl XlsxFillRectangle {
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
                            rect.b = string_to_int(&string_value);
                        }
                        b"l" => {
                            rect.l = string_to_int(&string_value);
                        }
                        b"r" => {
                            rect.r = string_to_int(&string_value);
                        }
                        b"t" => {
                            rect.t = string_to_int(&string_value);
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
