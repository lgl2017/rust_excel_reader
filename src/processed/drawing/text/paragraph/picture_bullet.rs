#[cfg(feature = "serde")]
use serde::Serialize;

use crate::common_types::HexColor;
use crate::packaging::relationship::XlsxRelationships;
use crate::processed::drawing::image::blip::Blip;
use crate::raw::drawing::scheme::color_scheme::XlsxColorScheme;
use crate::raw::drawing::text::paragraph::picture_bullet::XlsxPictureBullet;
use std::collections::BTreeMap;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picturebullet?view=openxml-3.0.1
///
/// This element specifies that a picture be applied to a set of bullets.
/// This element allows for any standard picture format graphic to be used instead of the typical bullet characters.
///
/// Example
/// ```
/// <a:pPr …>
///     <a:buBlip>
///         <a:blip r:embed="rId2"/>
///     </a:buBlip>
/// </a:pPr>
/// ```
// tag: buBlip
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct PictureBullet {
    pub blip: Blip,
}

impl PictureBullet {
    pub(crate) fn from_raw(
        raw: XlsxPictureBullet,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(blip) = Blip::from_raw(
            raw.blip.clone(),
            drawing_relationship,
            image_bytes,
            color_scheme.clone(),
            ref_color.clone(),
        ) else {
            return None;
        };

        return Some(Self { blip });
    }
}
