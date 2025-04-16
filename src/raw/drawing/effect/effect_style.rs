use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use crate::raw::drawing::{scene::scene_3d_type::Scene3DType, shape::shape_3d_type::Shape3DType};

use super::{effect_container::EffectDag, effect_list::EffectList};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectstyle?view=openxml-3.0.1
///
/// Example:
/// ```
/// <effectStyle>
///   <effectLst>
///     <outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
///       <schemeClr val="phClr">
///         <shade val="9000"/>
///         <satMod val="105000"/>
///         <alpha val="48000"/>
///       </schemeClr>
///     </outerShdw>
///   </effectLst>
/// </effectStyle>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct EffectStyle {
    /// Child Elements

    /// effectDag (Effect Container)	§20.1.8.25
    pub effect_dag: Option<EffectDag>,

    /// effectLst (Effect Container)	§20.1.8.26
    pub effect_lst: Option<EffectList>,

    /// scene3d (3D Scene Properties)	§20.1.4.1.26
    pub scene3d: Option<Scene3DType>,

    // sp3d (Apply 3D shape properties)
    pub shape3d: Option<Shape3DType>,
}

impl EffectStyle {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut style = Self {
            effect_dag: None,
            effect_lst: None,
            scene3d: None,
            shape3d: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    style.effect_dag = Some(EffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    style.effect_lst = Some(EffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scene3d" => {
                    style.scene3d = Some(Scene3DType::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp3d" => {
                    style.shape3d = Some(Shape3DType::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"effectStyle" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(style)
    }
}
