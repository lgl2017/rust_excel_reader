use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::drawing::{
        scheme::color_scheme::XlsxColorScheme,
        st_types::{st_percentage_to_float, st_text_point_to_pt},
        text::paragraph::paragraph_properties::XlsxParagraphProperties,
    },
};

use super::{
    auto_numbered_bullet::AutoNumberedBullet, character_bullet::CharacterBullet,
    picture_bullet::PictureBullet,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.defaultrunproperties?view=openxml-3.0.1
///
/// This element contains all default run level text properties for the text runs within a containing paragraph.
/// These properties are to be used when overriding properties have not been defined within the rPr element
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Bullet {
    /// bullet used in the paragraph
    pub r#type: BulletTypeValues,

    /// bullet color
    ///
    /// If no bullet color is specified along with this element, then the text color is used.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletcolor?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub color: Option<HexColor>,

    /// typeface of bullet font
    ///
    /// If no bullet font is specified along with this element then the paragraph font is used.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletfont?view=openxml-3.0.1
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub font: Option<String>,

    /// Bullet Size
    ///
    /// If no bullet size is specified along with this element then the size of the bullets for a paragraph should be of the same point size as the text run within which each bullet is contained.
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub size: Option<BulletSizeTypeValue>,
}

impl Bullet {
    pub(crate) fn from_paragraph_properties(
        raw: XlsxParagraphProperties,
        default: Option<Self>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        return Self {
            r#type: BulletTypeValues::from_paragraph_properties(
                raw.clone(),
                if let Some(d) = default.clone() {
                    Some(d.r#type)
                } else {
                    None
                },
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
            color: if let Some(c) = raw.clone().bullet_color {
                c.to_hex(color_scheme.clone(), None)
            } else if let Some(d) = default.clone() {
                d.color
            } else {
                None
            },
            font: if let Some(f) = raw.clone().bullet_font {
                f.typeface
            } else if let Some(d) = default.clone() {
                d.font
            } else {
                None
            },
            size: if let Some(s) = raw.clone().bullet_size_percentage {
                Some(BulletSizeTypeValue::Percentage(st_percentage_to_float(
                    s.val.unwrap_or(0) as i64,
                )))
            } else if let Some(s) = raw.clone().bullet_size_points {
                Some(BulletSizeTypeValue::Point(st_text_point_to_pt(
                    s.val.unwrap_or(0) as i64,
                )))
            } else if let Some(d) = default.clone() {
                d.size
            } else {
                None
            },
        };
    }
}

impl Bullet {
    pub(crate) fn default() -> Self {
        Self {
            r#type: BulletTypeValues::default(),
            color: None,
            font: None,
            size: None,
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BulletTypeValues {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nobullet?view=openxml-3.0.1
    ///
    /// This element specifies that the paragraph within which it is applied is to have no bullet formatting applied to it.
    NoBullet,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.characterbullet?view=openxml-3.0.1
    ///
    /// This element specifies that a character be applied to a set of bullets.
    /// These bullets are allowed to be any character in any font that the system is able to support.
    CharacterBullet(CharacterBullet),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.autonumberedbullet?view=openxml-3.0.1
    ///
    /// This element specifies that automatic numbered bullet points should be applied to a paragraph.
    /// These are not just numbers used as bullet points but instead automatically assigned numbers that are based on both buAutoNum attributes and paragraph level.
    AutoNumberedBullet(AutoNumberedBullet),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picturebullet?view=openxml-3.0.1
    ///
    /// This element specifies that a picture be applied to a set of bullets.
    /// This element allows for any standard picture format graphic to be used instead of the typical bullet characters.
    PictureBullet(PictureBullet),
}

impl BulletTypeValues {
    pub(crate) fn default() -> Self {
        Self::NoBullet
    }

    pub(crate) fn from_paragraph_properties(
        raw: XlsxParagraphProperties,
        default: Option<Self>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Self {
        if let Some(bullet) = raw.auto_numbered_bullet.clone() {
            return Self::AutoNumberedBullet(AutoNumberedBullet::from_raw(bullet));
        };
        if let Some(bullet) = raw.character_bullet.clone() {
            if let Some(bullet) = CharacterBullet::from_raw(bullet) {
                return Self::CharacterBullet(bullet);
            }
        };

        if let Some(bullet) = raw.picture_bullet.clone() {
            if let Some(bullet) = PictureBullet::from_raw(
                bullet,
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                None,
            ) {
                return Self::PictureBullet(bullet);
            }
        };

        if let Some(default) = default {
            return default;
        }

        return Self::default();
    }
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BulletSizeTypeValue {
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizepercentage?view=openxml-3.0.1
    ///
    /// This element specifies the size in percentage of the surrounding text to be used on bullet characters within a given paragraph.
    /// The size is specified using a percentage where 1000 is equal to 1 percent of the font size and 100000 is equal to 100 percent font of the font size.
    ///
    /// value range:
    /// - a minimum value of greater than or equal to 25000. (25%)
    /// - a maximum value of less than or equal to 400000. (400%)
    Percentage(f64),

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizepoints?view=openxml-3.0.1
    ///
    /// This element specifies the size in points to be used on bullet characters within a given paragraph.
    /// The size is specified using the points where 100 is equal to 1 point font and 1200 is equal to 12 point font.
    ///
    /// - a minimum value of greater than or equal to 100. (1pt)
    /// - a maximum value of less than or equal to 400000. (4000pt)
    Point(f64),
}
