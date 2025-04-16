use crate::raw::drawing::fill::FillStyleEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.filloverlay?view=openxml-3.0.1
/// specifies a fill overlay effect.
/// A fill overlay can be used to specify an additional fill for an object and blend the two fills together
pub type FillOverlay = FillStyleEnum;
