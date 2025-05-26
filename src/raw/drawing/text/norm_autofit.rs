use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STPercentage};

/// NormalAutoFit:
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.normalautofit?view=openxml-3.0.1
///
/// This element specifies that text within the text body should be normally auto-fit to the bounding box.
/// If this element is omitted, then noAutofit or auto-fit off is implied.
///
/// Example
/// ```
/// <a:bodyPr wrap="none" rtlCol="0">
///     <a:normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/>
/// </a:bodyPr>
/// ```
// normAutofit (Normal AutoFit)	§21.1.2.1.3
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNormAutoFit {
    // attributes
    /// Specifies the percentage of the original font size to which each run in the text body is scaled.
    ///
    /// In order to auto-fit text within a bounding box it is sometimes necessary to decrease the font size by a certain percentage.
    /// Using this attribute the font within a text box can be scaled based on the value provided.
    /// A value of 100% scales the text to 100%, while a value of 1% scales the text to 1%. If this attribute is omitted, then a value of 100% is implied.
    ///
    /// fontScale (Font Scale)
    pub font_scale: Option<STPercentage>,

    /// Specifies the percentage amount by which the line spacing of each paragraph in the text body is reduced.
    /// The reduction is applied by subtracting it from the original line spacing value.
    /// Using this attribute the vertical spacing between the lines of text can be scaled by a percent amount.
    /// A value of 100% reduces the line spacing by 100%, while a value of 1% reduces the line spacing by one percent.
    /// If this attribute is omitted, then a value of 0% is implied.
    ///
    /// lnSpcReduction (Line Space Reduction)
    pub ln_spc_reduction: Option<STPercentage>,
}

impl XlsxNormAutoFit {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fit = Self {
            font_scale: None,
            ln_spc_reduction: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"fontScale" => fit.font_scale = string_to_int(&string_value),
                        b"lnSpcReduction" => fit.ln_spc_reduction = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(fit)
    }
}
