use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::st_types::STPositiveCoordinate;
use crate::{
    helper::extract_val_attribute,
    raw::drawing::fill::{
        gradient_fill::XlsxGradientFill, no_fill::XlsxNoFill, pattern_fill::XlsxPatternFill,
        solid_fill::XlsxSolidFill,
    },
};

use super::{
    custom_dash::XlsxCustomDash, head_end::XlsxHeadEnd, line_join_bevel::XlsxLineJoinBevel,
    miter::XlsxMiter, round_line_join::XlsxRoundLineJoin, tail_end::XlsxTailEnd,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outline?view=openxml-3.0.1
///
/// This element specifies an outline style that can be applied to a number of different objects such as shapes and text.
/// The line allows for the specifying of many different types of outlines including even line dashes and bevels.
///
/// Example:
/// ```
/// <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
///     <a:solidFill>
///         <a:schemeClr val="phClr">
///             <a:shade val="95000" />
///             <a:satMod val="104999" />
///         </a:schemeClr>
///     </a:solidFill>
///     <a:prstDash val="solid" />
/// </a:ln>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxOutline {
    // child: extLst (Extension List)	§20.1.2.2.15 Not supported

    /* children */
    /// when present: specifies that an angle joint is used to connect lines
    pub bevel: Option<XlsxLineJoinBevel>,

    /// specifies a  custom dashing scheme
    // custDash (Custom Dash)	§20.1.8.21
    pub custom_dash: Option<XlsxCustomDash>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<XlsxGradientFill>,

    // headEnd (Line Head/End Style)	§20.1.8.38
    pub head_end: Option<XlsxHeadEnd>,

    // miter (Miter Line Join)	§20.1.8.43
    pub miter: Option<XlsxMiter>,

    /// No fill
    /// when present: indicates that the parent element is part of a group and should inherit the fill properties of the group.
    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<XlsxNoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<XlsxPatternFill>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetdash?view=openxml-3.0.1
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
    // tag: prstDash > attibute: val
    pub present_dash: Option<String>,

    /// when present: specifies that lines joined together have a round join
    // round (Round Line Join)	§20.1.8.52
    pub round: Option<XlsxRoundLineJoin>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<XlsxSolidFill>,

    // tailEnd (Tail line end style)
    pub tail_end: Option<XlsxTailEnd>,

    /* Attributes */
    /// Specifies the alignment to be used for the underline stroke.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.penalignmentvalues?view=openxml-3.0.1
    // algn (Stroke Alignment)
    pub alignment: Option<String>,

    /// Specifies the ending caps that should be used for this line such as rounded, flat, etc
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linecapvalues?view=openxml-3.0.1
    /// If this attribute is omitted, than a value of square is assumed.
    // cap (Line Ending Cap Type)
    pub cap: Option<String>,

    /// Specifies the compound line type to be used for the underline stroke.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.compoundlinevalues?view=openxml-3.0.1
    /// If this attribute is omitted, then a value of sng (single) is assumed.
    // cmpd (Compound Line Type)
    pub compound: Option<String>,

    /// Specifies the width to be used for the underline stroke.
    /// If this attribute is omitted, then a value of 0 is assumed.
    // w (Line Width)
    pub w: Option<STPositiveCoordinate>,
}

impl XlsxOutline {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut scheme = Self {
            // children
            bevel: None,
            custom_dash: None,
            gradient_fill: None,
            head_end: None,
            miter: None,
            no_fill: None,
            pattern_fill: None,
            present_dash: None,
            round: None,
            solid_fill: None,
            tail_end: None,
            // attributes
            alignment: None,
            cap: None,
            compound: None,
            w: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"algn" => {
                            scheme.alignment = Some(string_value);
                        }
                        b"cap" => {
                            scheme.cap = Some(string_value);
                        }
                        b"cmpd" => {
                            scheme.compound = Some(string_value);
                        }
                        b"w" => {
                            scheme.w = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bevel" => {
                    scheme.bevel = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"custDash" => {
                    scheme.custom_dash = Some(XlsxCustomDash::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    scheme.gradient_fill = Some(XlsxGradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"headEnd" => {
                    scheme.head_end = Some(XlsxHeadEnd::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"miter" => {
                    scheme.miter = Some(XlsxMiter::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    scheme.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    scheme.pattern_fill = Some(XlsxPatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstDash" => {
                    let val = extract_val_attribute(e)?;
                    scheme.present_dash = val;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"round" => {
                    scheme.round = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    scheme.solid_fill = XlsxSolidFill::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tailEnd" => {
                    scheme.tail_end = Some(XlsxTailEnd::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ln" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(scheme)
    }
}
