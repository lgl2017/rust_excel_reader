pub mod auto_numbered_bullet;
pub mod bullet;
pub mod character_bullet;
pub mod paragraph_properties;
pub mod picture_bullet;
pub mod spacing_type;
pub mod tab_alignment_values;
pub mod tab_stop;

#[cfg(feature = "serde")]
use serde::Serialize;

use std::collections::BTreeMap;

use super::{font::Font, run_type::RunTypeValues, text_run_properties::TextRunProperties};
use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme,
            text::paragraph::text_paragraphs::XlsxTextParagraphs,
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};
use paragraph_properties::ParagraphProperties;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.paragraph?view=openxml-3.0.1
///
/// This element specifies the presence of a paragraph of text within the containing text body.
/// The paragraph is the highest level text separation mechanism within a text body.
/// A paragraph can contain text paragraph properties associated with the paragraph.
/// If no properties are listed then properties specified in the defPPr element are used.
///
/// Example:
/// ```
/// <a:p>
///     <a:pPr marL="0" marR="0" indent="0" algn="l" defTabSz="457200" rtl="0"
///         fontAlgn="auto" latinLnBrk="0" hangingPunct="0">
///         <a:lnSpc>
///             <a:spcPct val="100000" />
///         </a:lnSpc>
///         <a:spcBef>
///             <a:spcPts val="0" />
///         </a:spcBef>
///         <a:spcAft>
///             <a:spcPts val="0" />
///         </a:spcAft>
///         <a:buClrTx />
///         <a:buSzTx />
///         <a:buFontTx />
///         <a:buNone />
///     </a:pPr>
///     <a:r>
///         <a:rPr kumimoji="0" lang="en-US" sz="1100" b="0" i="0" u="none"
///             strike="noStrike" cap="none" spc="0" normalizeH="0" baseline="0">
///             <a:ln>
///                 <a:noFill />
///             </a:ln>
///             <a:solidFill>
///                 <a:srgbClr val="000000" />
///             </a:solidFill>
///             <a:effectLst />
///             <a:uFillTx />
///             <a:latin typeface="+mn-lt" />
///             <a:ea typeface="+mn-ea" />
///             <a:cs typeface="+mn-cs" />
///             <a:sym typeface="Helvetica Neue" />
///         </a:rPr>
///         <a:t>Text</a:t>
///     </a:r>
///     <a:r>
///         <a:rPr lang="en-US" sz="1100" />
///         <a:t> Text</a:t>
///     </a:r>
/// </a:p>
/// ```
///
/// p (Text Paragraphs)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Paragraph {
    /// Paragraph Properties
    pub paragraph_properties: ParagraphProperties,

    /// Runs
    ///
    /// * TextRun
    /// * LineBreak
    /// * TextField (contains generated text that the application should update periodically)
    pub runs: Vec<RunTypeValues>,

    /// endParaRPr (End Paragraph Run Properties)
    ///
    /// This element specifies the text run properties that are to be used
    /// - if another run is inserted after the last run specified or
    /// - when an emtpy line (paragraph is inserted).
    #[cfg_attr(feature = "serde", serde(skip_serializing_if = "Option::is_none"))]
    pub end_paragraph_run_properties: Option<TextRunProperties>,
}

impl Paragraph {
    pub(crate) fn from_raw(
        raw: XlsxTextParagraphs,
        default_paragraph_properties: Option<ParagraphProperties>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Self {
        let paragraph_properties = if let Some(ppr) = raw.clone().paragraph_properties {
            Some(*ppr)
        } else {
            None
        };

        let default_run_properties = if let Some(p_pr) = paragraph_properties.clone() {
            Some(TextRunProperties::from_raw(
                p_pr.default_run_properties,
                None,
                drawing_relationship.clone(),
                defined_names.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
                ref_font.clone(),
                font_ref_color.clone(),
            ))
        } else {
            None
        };

        let runs: Vec<RunTypeValues> = raw
            .clone()
            .runs
            .unwrap_or(vec![])
            .into_iter()
            .map(|r| {
                RunTypeValues::from_raw(
                    r,
                    default_run_properties.clone(),
                    drawing_relationship.clone(),
                    defined_names.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_font.clone(),
                    font_ref_color.clone(),
                )
            })
            .collect();

        return Self {
            paragraph_properties: ParagraphProperties::from_raw(
                paragraph_properties.clone(),
                default_paragraph_properties.clone(),
                drawing_relationship.clone(),
                image_bytes.clone(),
                color_scheme.clone(),
            ),
            runs,
            end_paragraph_run_properties: if let Some(e_pr) =
                raw.clone().end_paragraph_run_properties
            {
                Some(TextRunProperties::from_raw(
                    Some(e_pr.clone()),
                    default_run_properties,
                    drawing_relationship.clone(),
                    defined_names.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_font.clone(),
                    font_ref_color.clone(),
                ))
            } else {
                None
            },
        };
    }
}
