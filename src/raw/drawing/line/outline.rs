use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::{
    helper::{extract_val_attribute, string_to_int},
    raw::drawing::fill::{
        gradient_fill::GradientFill, no_fill::NoFill, pattern_fill::PatternFill,
        solid_fill::SolidFill,
    },
};

use super::{
    custom_dash::CustomDash, head_end::HeadEnd, line_join_bevel::LineJoinBevel, miter::Miter,
    round_line_join::RoundLineJoin, tail_end::TailEnd,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.outline?view=openxml-3.0.1
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
pub struct Outline {
    // child: extLst (Extension List)	§20.1.2.2.15 Not supported

    /* children */
    /// when present: specifies that an angle joint is used to connect lines
    pub bevel: Option<LineJoinBevel>,

    /// specifies a  custom dashing scheme
    // custDash (Custom Dash)	§20.1.8.21
    pub custom_dash: Option<CustomDash>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<GradientFill>,

    // headEnd (Line Head/End Style)	§20.1.8.38
    pub head_end: Option<HeadEnd>,

    // miter (Miter Line Join)	§20.1.8.43
    pub miter: Option<Miter>,

    /// No fill
    /// when present: indicates that the parent element is part of a group and should inherit the fill properties of the group.
    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<NoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<PatternFill>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetdash?view=openxml-3.0.1
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetlinedashvalues?view=openxml-3.0.1
    // tag: prstDash > attibute: val
    pub present_dash: Option<String>,

    /// when present: specifies that lines joined together have a round join
    // round (Round Line Join)	§20.1.8.52
    pub round: Option<RoundLineJoin>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<SolidFill>,

    // tailEnd (Tail line end style)
    pub tail_end: Option<TailEnd>,

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
    pub w: Option<i64>,
}

impl Outline {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
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
            cap: Some("square".to_owned()),
            compound: Some("sng".to_owned()),
            w: Some(0),
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
                            scheme.w = string_to_int(&string_value);
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
                    scheme.custom_dash = Some(CustomDash::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    scheme.gradient_fill = Some(GradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"headEnd" => {
                    scheme.head_end = Some(HeadEnd::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"miter" => {
                    scheme.miter = Some(Miter::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    scheme.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    scheme.pattern_fill = Some(PatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstDash" => {
                    let val = extract_val_attribute(e)?;
                    scheme.present_dash = val;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"round" => {
                    scheme.round = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    scheme.solid_fill = SolidFill::load(reader, e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tailEnd" => {
                    scheme.tail_end = Some(TailEnd::load(e)?);
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
