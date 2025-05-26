use std::collections::BTreeMap;

#[cfg(feature = "serde")]
use serde::Serialize;

use crate::{
    common_types::HexColor,
    packaging::relationship::XlsxRelationships,
    raw::{
        drawing::{
            scheme::color_scheme::XlsxColorScheme, text::shape_text_body::XlsxShapeTextBody,
        },
        spreadsheet::workbook::defined_name::XlsxDefinedNames,
    },
};

use super::{
    body_properties::BodyProperties,
    font::Font,
    paragraph::{paragraph_properties::ParagraphProperties, Paragraph},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textbody?view=openxml-3.0.1
///
/// This element specifies the existence of text to be contained within the corresponding shape.
/// All visible text and visible text related properties are contained within this element.
/// There can be multiple paragraphs and within paragraphs multiple runs of text.
///
/// Example:
/// ```
/// <xdr:txBody>
/// <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip"
///     vert="horz" wrap="square" lIns="50800" tIns="50800" rIns="50800"
///     bIns="50800" numCol="1" spcCol="38100" rtlCol="0" anchor="t">
///     <a:spAutoFit />
/// </a:bodyPr>
/// <a:lstStyle />
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
/// </a:p>
/// <a:p>
///     <a:r>
///         <a:rPr lang="en-US" sz="1100" />
///         <a:t>text box</a:t>
///     </a:r>
/// </a:p>
/// <a:p>
///     <a:pPr marL="228600" indent="-228600">
///         <a:buFont typeface="+mj-lt" />
///         <a:buAutoNum type="arabicPeriod" />
///     </a:pPr>
///     <a:endParaRPr lang="en-US" sz="1100" />
/// </a:p>
/// </xdr:txBody>
/// ```
///
/// txBody (Shape Text Body)
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct ShapeTextBody {
    /// Body Properties
    ///
    /// defines the body properties for the text body within a shape.
    pub body_properties: BodyProperties,

    // Paragraphs
    pub text_paragraph: Vec<Paragraph>,
}

impl ShapeTextBody {
    pub(crate) fn from_raw(
        raw: Option<XlsxShapeTextBody>,
        drawing_relationship: XlsxRelationships,
        defined_names: XlsxDefinedNames,
        image_bytes: BTreeMap<String, Vec<u8>>,
        color_scheme: Option<XlsxColorScheme>,
        ref_font: Option<Font>,
        font_ref_color: Option<HexColor>,
    ) -> Option<Self> {
        let Some(raw) = raw else { return None };
        let mut default_paragraph_properties: Option<ParagraphProperties> = None;

        if let Some(lst) = raw.clone().text_list_style {
            if let Some(def_pr) = lst.default_paragraph_style {
                default_paragraph_properties = Some(ParagraphProperties::from_raw(
                    Some(*def_pr.clone()),
                    None,
                    drawing_relationship.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                ))
            }
        }

        let paragraphs: Vec<Paragraph> = raw
            .clone()
            .text_paragraph
            .unwrap_or(vec![])
            .into_iter()
            .map(|p| {
                Paragraph::from_raw(
                    p,
                    default_paragraph_properties.clone(),
                    drawing_relationship.clone(),
                    defined_names.clone(),
                    image_bytes.clone(),
                    color_scheme.clone(),
                    ref_font.clone(),
                    font_ref_color.clone(),
                )
            })
            .collect();

        return Some(Self {
            body_properties: BodyProperties::from_raw(
                raw.clone().body_properties,
                color_scheme.clone(),
            ),
            text_paragraph: paragraphs,
        });
    }
}
