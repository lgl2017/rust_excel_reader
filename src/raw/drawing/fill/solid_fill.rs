use crate::raw::drawing::color::ColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.solidfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:solidFill>
///     <a:schemeClr val="phClr" />
/// </a:solidFill>
/// ```
// tag: solidFill
pub type SolidFill = ColorEnum;
