use crate::raw::drawing::line::outline::Outline;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.underline?view=openxml-3.0.1
///
/// Secifies the properties for the stroke of the underline that is present within a run of text.
///
/// Example
/// ```
/// <a:rPr …>
///     <a:uLn algn="r">
/// </a:rPr>
/// ```
pub type Underline = Outline;
