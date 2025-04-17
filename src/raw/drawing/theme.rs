use std::io::{Read, Seek};

use anyhow::bail;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::excel::xml_reader;

use super::{
    color::custom_color::{load_custom_color_list, CustomColorList},
    default::object_defaults::ObjectDefaults,
    scheme::extra_color_scheme::{load_extra_color_scheme_list, ExtraColorSchemeList},
    theme_elements::ThemeElements,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1
///
/// Root element of DrawingML Theme part
///
/// theme (Theme)
#[derive(Debug, Clone, PartialEq)]
pub struct Theme {
    // child: extLst (Extension List)	§20.1.2.2.15 Not supported

    // children
    // tag: custClrLst
    pub custom_color_list: Option<Box<CustomColorList>>,

    // tag: extraClrSchemeLst
    pub extra_color_scheme_list: Option<Box<ExtraColorSchemeList>>,

    // objectDefaults (Object Defaults)	§20.1.6.7
    pub object_defaults: Option<Box<ObjectDefaults>>,

    // tag: themeElements
    pub theme_elements: Option<Box<ThemeElements>>,

    // attributes
    pub name: Option<String>,
}

impl Theme {
    pub(crate) fn load(
        zip: &mut ZipArchive<impl Read + Seek>,
        path: Vec<String>,
    ) -> anyhow::Result<Self> {
        let mut theme = Self {
            name: None,
            custom_color_list: None,
            extra_color_scheme_list: None,
            theme_elements: None,
            object_defaults: None,
        };

        let Some(path) = path.first() else {
            return Ok(theme);
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(theme);
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"theme" => {
                    let attributes = e.attributes();
                    for a in attributes {
                        match a {
                            Ok(a) => {
                                let string_value = String::from_utf8(a.value.to_vec())?;
                                match a.key.local_name().as_ref() {
                                    b"name" => {
                                        theme.name = Some(string_value);
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
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"custClrLst" => {
                    let colors = load_custom_color_list(&mut reader)?;
                    theme.custom_color_list = Some(Box::new(colors));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extraClrSchemeLst" => {
                    let schemes = load_extra_color_scheme_list(&mut reader)?;
                    theme.extra_color_scheme_list = Some(Box::new(schemes));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"objectDefaults" => {
                    let defaults: ObjectDefaults = ObjectDefaults::load(&mut reader)?;
                    theme.object_defaults = Some(Box::new(defaults));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"themeElements" => {
                    let theme_elements = ThemeElements::load(&mut reader)?;
                    theme.theme_elements = Some(Box::new(theme_elements));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"theme" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(theme);
    }
}
