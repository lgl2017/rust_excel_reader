/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefilltext?view=openxml-3.0.1
///
/// This element specifies that the fill color of an underline for a run of text should be of the same color as the text run within which it is contained.
///
/// Example:
/// ```
/// <a:rPr …>
///     <a:uFillTx />
/// </a:rPr>
/// ```
// tag: buFontTx
pub type UnderlineFillText = bool;
