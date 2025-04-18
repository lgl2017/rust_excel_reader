use crate::raw::drawing::fill::XlsxFillStyleEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefill?view=openxml-3.0.1
///
/// This element specifies the fill color of an underline for a run of text.
///
/// Example:
/// ```
/// <a:rPr …>
///     <a:uFill>
///         <a:solidFill>
///             <a:srgbClr val="FFFF00"/>
///         </a:solidFill>
///     </a:uFill>
/// </a:rPr>
/// ```
pub type XlsxUnderlineFill = XlsxFillStyleEnum;
