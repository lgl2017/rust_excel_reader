use crate::excel::XmlReader;

use crate::raw::drawing::color::ColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1#properties
/// This element specifies a duotone effect.
/// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
// tag: duotone
#[derive(Debug, Clone, PartialEq)]
pub struct Duotone {
    colors: Option<Vec<ColorEnum>>,
}

impl Duotone {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let colors: Vec<ColorEnum> = ColorEnum::load_list(reader, b"duotone")?;

        Ok(Self {
            colors: Some(colors),
        })
    }
}
