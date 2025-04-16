use crate::raw::drawing::effect::bevel::Bevel;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelbottom?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d>
///     <a:bevelB w="139700" h="127000" prst="coolSlant"/>
/// </a:sp3d>
/// ```
pub type BevelBottom = Bevel;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.beveltop?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d>
///     <a:bevelT w="139700" h="127000" prst="coolSlant"/>
/// </a:sp3d>
/// ```
pub type BevelTop = Bevel;
