use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::{backdrop::BackDrop, camera::Camera, light_rig::LightRig};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.scene3dtype?view=openxml-3.0.1
///
/// Example
/// ```
/// <a:scene3d>
///     <a:backdrop>
///         <anchor x="123" y="23" z="10000"/>
///         <norm dx="123" dy="23" dz="10000"/>
///         <up dx="123" dy="23" dz="10000"/>
///     </a:backdrop>
///     <a:camera prst="orthographicFront">
///         <a:rot lat="19902513" lon="17826689" rev="1362739"/>
///     </a:camera>
///     <a:lightRig rig="twoPt" dir="t">
///         <a:rot lat="0" lon="0" rev="6000000"/>
///     </a:lightRig>
/// </<a:scene3d>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Scene3DType {
    // extLst (Extension List)	§20.1.2.2.15 Not supported

    // children
    // backdrop (Backdrop Plane)	§20.1.5.2
    pub backdrop: Option<BackDrop>,

    // camera (Camera)	§20.1.5.5
    pub camera: Option<Camera>,

    // lightRig (Light Rig)
    pub light_rig: Option<LightRig>,
}

impl Scene3DType {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut scene3d = Self {
            backdrop: None,
            camera: None,
            light_rig: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"backdrop" => {
                    scene3d.backdrop = Some(BackDrop::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"camera" => {
                    scene3d.camera = Some(Camera::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lightRig" => {
                    scene3d.light_rig = Some(LightRig::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"scene3d" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(scene3d)
    }
}
