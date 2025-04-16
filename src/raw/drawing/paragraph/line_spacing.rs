use super::spacing::SpacingEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linespacing?view=openxml-3.0.1
///
/// This element specifies the vertical line spacing that is to be used within a paragraph.
/// This can be specified in two different ways, percentage spacing and font point spacing.
///
/// Example:
/// ```
/// <a:pPr>
///     <a:lnSpc>
///         <a:spcPct val="200%"/>
///     </a:lnSpc>
/// </a:pPr>
/// ```
// tag: lnSpc
pub type LineSpacing = SpacingEnum;
