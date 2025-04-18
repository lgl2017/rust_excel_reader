use super::spacing::XlsxSpacingEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spaceafter?view=openxml-3.0.1
///
/// This element specifies the amount of vertical white space that is present after a paragraph.
/// This can be specified in two different ways, percentage spacing and font point spacing.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:spcBef>
///         <a:spcPts val="1800"/>
///     </a:spcBef>
///     <a:spcAft>
///         <a:spcPts val="600"/>
///     </a:spcAft>
/// </a:pPr>
/// ```
// tag: spcAft
pub type XlsxSpaceAfter = XlsxSpacingEnum;
