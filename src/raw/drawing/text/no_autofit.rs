/// NoAutoFit:  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.noautofit?view=openxml-3.0.1
/// This element specifies that text within the text body should not be auto-fit to the bounding box.
/// Auto-fitting is when text within a text box is scaled in order to remain inside the text box.
///
/// Example
/// ```
/// <a:bodyPr wrap="none" rtlCol="0">
///     <a:noAutofit/>
/// </a:bodyPr>
/// ```
// noAutofit (No AutoFit)	§21.1.2.1.2
pub type NoAutoFit = bool;
