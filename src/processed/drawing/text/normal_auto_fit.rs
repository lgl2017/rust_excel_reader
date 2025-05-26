#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{st_types::st_percentage_to_float, text::norm_autofit::XlsxNormAutoFit};

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
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct NormalAutoFit {
    // attributes
    /// Specifies the percentage of the original font size to which each run in the text body is scaled.
    ///
    /// In order to auto-fit text within a bounding box it is sometimes necessary to decrease the font size by a certain percentage.
    /// Using this attribute the font within a text box can be scaled based on the value provided.
    ///
    /// A value of 100% scales the text to 100%, while a value of 1% scales the text to 1%. If this attribute is omitted, then a value of 100% is implied.
    ///
    /// fontScale (Font Scale)
    pub font_scale: f64,

    /// Specifies the percentage amount by which the line spacing of each paragraph in the text body is reduced.
    ///
    /// The reduction is applied by subtracting it from the original line spacing value.
    /// Using this attribute the vertical spacing between the lines of text can be scaled by a percent amount.
    /// A value of 100% reduces the line spacing by 100%, while a value of 1% reduces the line spacing by one percent.
    ///
    /// If this attribute is omitted, then a value of 0% is implied.
    ///
    /// lnSpcReduction (Line Space Reduction)
    pub line_space_reduction: f64,
}

impl NormalAutoFit {
    pub(crate) fn from_raw(raw: XlsxNormAutoFit) -> Self {
        let scale = if let Some(font_scale) = raw.font_scale.clone() {
            st_percentage_to_float(font_scale)
        } else {
            1.0
        };
        let reduction = if let Some(ln_spc_reduction) = raw.ln_spc_reduction.clone() {
            st_percentage_to_float(ln_spc_reduction)
        } else {
            0.0
        };
        return Self {
            font_scale: scale,
            line_space_reduction: reduction,
        };
    }
}
