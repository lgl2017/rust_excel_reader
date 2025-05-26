use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1
///
/// This element specifies a duotone effect.
/// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
// tag: duotone
pub type XlsxDuotone = XlsxColorEnum;
