use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    raw::drawing::{
        effect::{effect_container::XlsxEffectDag, effect_list::XlsxEffectList},
        fill::{
            blip_fill::XlsxBlipFill, gradient_fill::XlsxGradientFill, group_fill::XlsxGroupFill,
            no_fill::XlsxNoFill, pattern_fill::XlsxPatternFill, solid_fill::XlsxSolidFill,
        },
        scene::scene_3d_type::XlsxScene3DType,
    },
};

use super::transform_group::XlsxTransformGroup;

/// - VisualGroupShapeProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.visualgroupshapeproperties?view=openxml-3.0.1
/// - SpreadSheet.GroupShapeProperties: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.groupshapeproperties?view=openxml-3.0.1
///
/// This element specifies the properties that are to be common across all of the shapes within the corresponding group.
/// If there are any conflicting properties within the group shape properties and the individual shape properties then the individual shape properties should take precedence.
///
/// Example:
///
/// ```
/// <p:grpSpPr>
///     <a:xfrm>
///         <a:off x="838200" y="990600"/>
///         <a:ext cx="2426208" cy="978408"/>
///         <a:chOff x="838200" y="990600"/>
///         <a:chExt cx="2426208" cy="978408"/>
///     </a:xfrm>
/// </p:grpSpPr>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxVisualGroupShapeProperties {
    // Child Elements
    // extLst (Extension List) Not supported

    // blipFill (Picture Fill)	§20.1.8.14
    pub blip_fill: Option<XlsxBlipFill>,

    // effectDag (Effect Container)	§20.1.8.25
    pub effect_dag: Option<XlsxEffectDag>,

    // effectLst (Effect Container)	§20.1.8.26
    pub effect_list: Option<XlsxEffectList>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<XlsxGradientFill>,

    // grpFill (Group Fill)	§20.1.8.35
    pub group_fill: Option<XlsxGroupFill>,

    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<XlsxNoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<XlsxPatternFill>,

    // scene3d (3D Scene Properties)	§20.1.4.1.26
    pub scene3d: Option<XlsxScene3DType>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<XlsxSolidFill>,

    // xfrm (2D Transform for Grouped Objects)	§20.1.7.5
    pub transform2d: Option<XlsxTransformGroup>,

    // Attributes
    /// Specifies that the picture should be rendered using only black and white coloring.
    /// That is the coloring information for the picture should be converted to either black or white when rendering the picture.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blackwhitemodevalues?view=openxml-3.0.1
    ///
    /// bwMode (Black and White Mode)
    pub black_white_mode: Option<String>,
}

impl XlsxVisualGroupShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut properties = Self {
            blip_fill: None,
            effect_dag: None,
            effect_list: None,
            gradient_fill: None,
            group_fill: None,
            no_fill: None,
            pattern_fill: None,
            scene3d: None,
            solid_fill: None,
            transform2d: None,
            // attributes
            black_white_mode: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"bwMode" => {
                            properties.black_white_mode = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blipFill" => {
                    properties.blip_fill = Some(XlsxBlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    properties.effect_dag = Some(XlsxEffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    properties.effect_list = Some(XlsxEffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    properties.gradient_fill = Some(XlsxGradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpFill" => {
                    properties.group_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    properties.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    properties.pattern_fill = Some(XlsxPatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scene3d" => {
                    properties.scene3d = Some(XlsxScene3DType::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    properties.solid_fill = XlsxSolidFill::load(reader, b"solidFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    properties.transform2d = Some(XlsxTransformGroup::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"grpSpPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(properties)
    }
}
