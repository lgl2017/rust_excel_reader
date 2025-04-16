use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::{
    effect::{effect_container::EffectDag, effect_list::EffectList},
    fill::{
        blip_fill::BlipFill, gradient_fill::GradientFill, group_fill::GroupFill, no_fill::NoFill,
        pattern_fill::PatternFill, solid_fill::SolidFill,
    },
    line::outline::Outline,
    scene::scene_3d_type::Scene3DType,
};

use super::{
    custom_geometry::CustomGeometry, preset_geometry::PresetGeometry, shape_3d_type::Shape3DType,
    transform_2d::Transform2D,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeproperties?view=openxml-3.0.1
///
/// This element specifies the visual shape properties that can be applied to a shape.
///
/// Example
/// ```
/// <a:spPr>
///     <a:noFill />
///     <a:ln w="12700" cap="flat">
///         <a:solidFill>
///             <a:srgbClr val="000000" />
///         </a:solidFill>
///         <a:prstDash val="solid" />
///         <a:miter lim="400000" />
///     </a:ln>
///     <a:effectLst />
///     <a:sp3d />
/// </a:spPr>
/// ```
// tag: spPr
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeProperties {
    // extLst (Extension List)	Not Supported

    // Child Elements	Subclause
    // blipFill (Picture Fill)	§20.1.8.14
    pub blip_fill: Option<BlipFill>,

    // custGeom (Custom Geometry)	§20.1.9.8
    pub custom_geometry: Option<CustomGeometry>,

    // effectDag (Effect Container)	§20.1.8.25
    pub effect_dag: Option<EffectDag>,

    // effectLst (Effect Container)	§20.1.8.26
    pub effect_list: Option<EffectList>,

    // gradFill (Gradient Fill)	§20.1.8.33
    pub gradient_fill: Option<GradientFill>,

    // grpFill (Group Fill)	§20.1.8.35
    pub group_fill: Option<GroupFill>,

    // ln (Outline)	§20.1.2.2.24
    pub outline: Option<Outline>,

    // noFill (No Fill)	§20.1.8.44
    pub no_fill: Option<NoFill>,

    // pattFill (Pattern Fill)	§20.1.8.47
    pub pattern_fill: Option<PatternFill>,

    // prstGeom (Preset geometry)	§20.1.9.18
    pub preset_gemoetry: Option<PresetGeometry>,

    // scene3d (3D Scene Properties)	§20.1.4.1.26
    pub scene3d: Option<Scene3DType>,

    // solidFill (Solid Fill)	§20.1.8.54
    pub solid_fill: Option<SolidFill>,

    // sp3d (Apply 3D shape properties)	§20.1.5.12
    pub shape3d: Option<Shape3DType>,

    // xfrm (2D Transform for Individual Objects)
    pub transform2d: Option<Transform2D>,

    // Attributes
    /// Specifies that the picture should be rendered using only black and white coloring.
    /// That is the coloring information for the picture should be converted to either black or white when rendering the picture.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blackwhitemodevalues?view=openxml-3.0.1
    // bwMode (Black and White Mode)
    pub black_white_mode: Option<String>,
}

impl ShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut properties = Self {
            blip_fill: None,
            custom_geometry: None,
            effect_dag: None,
            effect_list: None,
            gradient_fill: None,
            group_fill: None,
            outline: None,
            no_fill: None,
            pattern_fill: None,
            preset_gemoetry: None,
            scene3d: None,
            solid_fill: None,
            shape3d: None,
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
                    properties.blip_fill = Some(BlipFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"custGeom" => {
                    properties.custom_geometry = Some(CustomGeometry::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectDag" => {
                    properties.effect_dag = Some(EffectDag::load_effect_dag(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectLst" => {
                    properties.effect_list = Some(EffectList::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradFill" => {
                    properties.gradient_fill = Some(GradientFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpFill" => {
                    properties.group_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ln" => {
                    properties.outline = Some(Outline::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"noFill" => {
                    properties.no_fill = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pattFill" => {
                    properties.pattern_fill = Some(PatternFill::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"prstGeom" => {
                    properties.preset_gemoetry = Some(PresetGeometry::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scene3d" => {
                    properties.scene3d = Some(Scene3DType::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"solidFill" => {
                    properties.solid_fill = SolidFill::load(reader, b"solidFill")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp3d" => {
                    properties.shape3d = Some(Shape3DType::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    properties.transform2d = Some(Transform2D::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"spPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(properties)
    }
}
