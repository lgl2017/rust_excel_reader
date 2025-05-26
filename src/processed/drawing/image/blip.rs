use std::{collections::BTreeMap, path::PathBuf, str::FromStr};

use crate::{
    common_types::HexColor,
    packaging::relationship::{zip_path_for_id, XlsxRelationships, EXTERNAL_TARGET_MODE},
    processed::drawing::effect::effect_container::EffectContainer,
    raw::drawing::{image::blip::XlsxBlip, scheme::color_scheme::XlsxColorScheme},
};

#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip?view=openxml-3.0.1
///
/// This element specifies the existence of an image (binary large image or picture) and contains a reference to the image data.
///
/// Example
/// ```
/// <a:blip
///     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///     r:embed="rId1" />
/// ```
#[cfg_attr(feature = "serde", derive(Serialize))]
#[derive(Debug, PartialEq, Clone)]
pub struct Blip {
    pub source: BlipSourceType,

    pub compression_state: BlipCompressionStateType,

    pub effects: Box<EffectContainer>,
}

impl Blip {
    pub(crate) fn from_raw(
        raw: Option<XlsxBlip>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };
        let Some(source) = BlipSourceType::from_raw(
            Some(raw.clone()),
            drawing_relationship.clone(),
            image_bytes.clone(),
        ) else {
            return None;
        };

        return Some(Self {
            effects: EffectContainer::from_raw_blip(
                raw.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_color.clone(),
            ),
            source,
            compression_state: BlipCompressionStateType::from_string(raw.cstate),
        });
    }
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BlipSourceType {
    /// image data resides locally within the excel
    Internal(InternalSource),

    /// external link to the picture
    External(String),
}

impl BlipSourceType {
    pub(crate) fn from_raw(
        raw_blip: Option<XlsxBlip>,
        drawing_relationship: XlsxRelationships,
        image_bytes: BTreeMap<String, Vec<u8>>,
    ) -> Option<Self> {
        let Some(raw_blip) = raw_blip else {
            return None;
        };
        if let Some(embed) = raw_blip.clone().embed {
            return Self::parse_rel_helper(&embed, drawing_relationship.clone(), image_bytes);
        }
        if let Some(link) = raw_blip.clone().link {
            if let Some(source) =
                Self::parse_rel_helper(&link, drawing_relationship.clone(), image_bytes)
            {
                return Some(source);
            } else {
                return Some(Self::External(link));
            }
        }
        return None;
    }

    fn parse_rel_helper(
        r_id: &str,
        drawing_relationships: XlsxRelationships,
        // get bytes for a rel_id
        image_bytes: BTreeMap<String, Vec<u8>>,
    ) -> Option<Self> {
        let rel: XlsxRelationships = drawing_relationships
            .clone()
            .into_iter()
            .filter(|r| r.id == r_id.to_string())
            .collect();

        if let Some(first) = rel.first() {
            if first.target_mode == Some(EXTERNAL_TARGET_MODE.to_string()) {
                return Some(Self::External(first.target.clone()));
            }
            if let (Some(path), Some(bytes)) = (
                zip_path_for_id(&drawing_relationships, &first.id),
                image_bytes.get(&first.id),
            ) {
                let mut name = path.clone();
                match PathBuf::from_str(&path) {
                    Ok(file_path) => {
                        if let Some(n) = file_path.file_name() {
                            if let Some(s) = n.to_str() {
                                name = s.to_owned();
                            }
                        }
                    }
                    _ => (),
                };

                return Some(Self::Internal(InternalSource {
                    name,
                    bytes: bytes.to_owned(),
                }));
            }
        }
        return None;
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blipcompressionvalues?view=openxml-3.0.1
///
/// * Email
/// * HighQualityPrint
/// * None
/// * Print
/// * Screen
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum BlipCompressionStateType {
    Email,
    HighQualityPrint,
    None,
    Print,
    Screen,
}

impl BlipCompressionStateType {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::None };
        return match s.as_ref() {
            "email" => Self::Email,
            "hqprint" => Self::HighQualityPrint,
            "none" => Self::None,
            "print" => Self::Print,
            "screen" => Self::Screen,
            _ => Self::None,
        };
    }
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct InternalSource {
    pub name: String,

    #[cfg_attr(feature = "serde", serde(skip_serializing))]
    pub bytes: Vec<u8>,
}
