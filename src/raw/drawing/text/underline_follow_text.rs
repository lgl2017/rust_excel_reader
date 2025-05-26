/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underlinefollowstext?view=openxml-3.0.1
///
/// This element specifies that the stroke style of an underline for a run of text should be of the same as the text run within which it is contained.
///
/// Example:
/// ```
/// <a:rPr …>
///     <a:uLnTx />
/// </a:rPr>
/// ```
// tag: buFontTx
pub type XlsxUnderlineFollowsText = bool;
