use crate::raw::drawing::shape::rectangle::Rectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tilerectangle?view=openxml-3.0.1
pub type TileRectangle = Rectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.sourcerectangle?view=openxml-3.0.1
pub type SourceRectangle = Rectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filltorectangle?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
/// ```
pub type FillToRectangle = Rectangle;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillrectangle?view=openxml-3.0.1
pub type FillRectangle = Rectangle;
