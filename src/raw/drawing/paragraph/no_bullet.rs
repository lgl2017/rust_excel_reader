/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nobullet?view=openxml-3.0.1
///
/// This element specifies that the paragraph within which it is applied is to have no bullet formatting applied to it.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buNone/>
/// </a:pPr>
/// ```
pub type XlsxNoBullet = bool;
