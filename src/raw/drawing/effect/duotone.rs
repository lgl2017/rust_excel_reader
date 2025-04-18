use crate::excel::XmlReader;

use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.duotone?view=openxml-3.0.1#properties
/// This element specifies a duotone effect.
/// For each pixel, combines clr1 and clr2 through a linear interpolation to determine the new color for that pixel.
// tag: duotone
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDuotone {
    colors: Option<Vec<XlsxColorEnum>>,
}

impl XlsxDuotone {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let colors: Vec<XlsxColorEnum> = XlsxColorEnum::load_list(reader, b"duotone")?;

        Ok(Self {
            colors: Some(colors),
        })
    }
}
