use anyhow::bail;
use quick_xml::events::Event;
use std::io::{Read, Seek};
use zip::ZipArchive;

use crate::excel::xml_reader;

use super::{
    color::custom_color::{load_custom_color_list, XlsxCustomColorList},
    default::object_defaults::XlsxObjectDefaults,
    scheme::extra_color_scheme::{load_extra_color_scheme_list, XlsxExtraColorSchemeList},
    theme_elements::XlsxThemeElements,
};

#[cfg(feature = "drawing")]
use super::{
    effect::{effect_reference::XlsxEffectReference, effect_style::XlsxEffectStyle},
    fill::{fill_reference::XlsxFillReference, XlsxFillStyleEnum},
    line::{line_reference::XlsxLineReference, outline::XlsxOutline},
    scheme::format_scheme::XlsxFormatScheme,
    text::font::{font_reference::XlsxFontReference, XlsxFontBase},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1
///
/// Root element of DrawingML Theme part
///
/// theme (Theme)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTheme {
    // child: extLst (Extension List)	§20.1.2.2.15 Not supported

    // children
    // tag: custClrLst
    pub custom_color_list: Option<Box<XlsxCustomColorList>>,

    // tag: extraClrSchemeLst
    pub extra_color_scheme_list: Option<Box<XlsxExtraColorSchemeList>>,

    // objectDefaults (Object Defaults)	§20.1.6.7
    pub object_defaults: Option<Box<XlsxObjectDefaults>>,

    // tag: themeElements
    pub theme_elements: Option<Box<XlsxThemeElements>>,

    // attributes
    pub name: Option<String>,
}

impl XlsxTheme {
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
                    let defaults: XlsxObjectDefaults = XlsxObjectDefaults::load(&mut reader)?;
                    theme.object_defaults = Some(Box::new(defaults));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"themeElements" => {
                    let theme_elements = XlsxThemeElements::load(&mut reader)?;
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

#[cfg(feature = "drawing")]
impl XlsxTheme {
    fn get_format_scheme(&self) -> Option<XlsxFormatScheme> {
        let Some(theme_elements) = self.theme_elements.clone() else {
            return None;
        };
        return theme_elements.format_scheme.clone();
    }

    fn u64_to_usize(u64: u64) -> anyhow::Result<usize> {
        return Ok(TryInto::<usize>::try_into(u64)?);
    }

    /// Get outline reference by a `lnRef`
    ///
    /// Example:
    /// ```
    /// <lnRef idx="1">
    ///     <schemeClr val="accent2"/>
    /// </lnRef>
    /// ```
    pub(crate) fn get_line_from_ref(
        &self,
        reference: Option<XlsxLineReference>,
    ) -> Option<XlsxOutline> {
        let Some(reference) = reference else {
            return None;
        };
        let Some(index) = reference.index else {
            return None;
        };

        let Some(format_scheme) = self.get_format_scheme() else {
            return None;
        };

        let Some(line_style_list) = format_scheme.line_style_lst else {
            return None;
        };

        let Ok(index) = Self::u64_to_usize(index) else {
            return None;
        };

        // let Some(index) = index.checked_sub(1) else {
        //     return None;
        // };

        if index > line_style_list.len() - 1 {
            return None;
        };

        return Some(line_style_list[index].clone());
    }

    /// Get outline reference by a `fillRef`
    ///
    /// idx:
    /// - value of 0 or 1000 indicates no background,
    /// - values 1-999 refer to the index of a fill style within the fillStyleLst element,
    /// - values 1001 and above refer to the index of a background fill style within the bgFillStyleLst element.
    ///     For example: The value 1001 corresponds to the first background fill style, 1002 to the second background fill style, and so on.
    ///
    /// Example:
    /// ```
    /// <fillRef idx="0">
    ///     <schemeClr val="accent2"/>
    /// </fillRef>
    /// ```
    pub(crate) fn get_fill_from_ref(
        &self,
        reference: Option<XlsxFillReference>,
    ) -> Option<XlsxFillStyleEnum> {
        let Some(reference) = reference else {
            return None;
        };
        let Some(index) = reference.index else {
            return None;
        };

        let Some(format_scheme) = self.get_format_scheme() else {
            return None;
        };

        let Ok(index) = Self::u64_to_usize(index) else {
            return None;
        };

        // value of 0 or 1000 indicates no background,
        if index == 0 || index == 1000 {
            return Some(XlsxFillStyleEnum::NoFill(true));
        }

        // values 1-999 refer to the index of a fill style within the fillStyleLst element
        if (1..=999).contains(&index) {
            let Some(style_list) = format_scheme.fill_style_lst else {
                return None;
            };

            let Some(index) = index.checked_sub(1) else {
                return None;
            };

            if index > style_list.len() - 1 {
                return None;
            };

            return Some(style_list[index].clone());
        }

        // values 1001 and above refer to the index of a background fill style within the bgFillStyleLst element.
        // For example: The value 1001 corresponds to the first background fill style, 1002 to the second background fill style, and so on.
        let Some(index) = index.checked_sub(1001) else {
            return None;
        };

        let Some(style_list) = format_scheme.bg_fill_style_lst else {
            return None;
        };

        if index > style_list.len() - 1 {
            return None;
        };

        return Some(style_list[index].clone());
    }

    /// Get effect style reference by a `effectRef`
    ///
    /// Example:
    /// ```
    /// <effectRef idx="0">
    ///     <schemeClr val="accent2"/>
    /// </effectRef>
    /// ```
    pub(crate) fn get_effect_from_ref(
        &self,
        reference: Option<XlsxEffectReference>,
    ) -> Option<XlsxEffectStyle> {
        let Some(reference) = reference else {
            return None;
        };
        let Some(index) = reference.index else {
            return None;
        };

        let Some(format_scheme) = self.get_format_scheme() else {
            return None;
        };

        let Some(style_list) = format_scheme.effect_style_lst else {
            return None;
        };

        let Ok(index) = Self::u64_to_usize(index) else {
            return None;
        };

        if index > style_list.len() - 1 {
            return None;
        };

        return Some(style_list[index].clone());
    }

    /// Get font reference by a `fontRef`
    ///
    /// idx:
    /// - `major`
    /// - `minor`
    /// - `none`
    ///
    /// Example:
    /// ```
    /// <fontRef idx="minor">
    ///     <schemeClr val="tx1"/>
    /// </fontRef>
    /// ```
    pub(crate) fn get_font_from_ref(
        &self,
        reference: Option<XlsxFontReference>,
    ) -> Option<XlsxFontBase> {
        let Some(reference) = reference else {
            return None;
        };
        let Some(index) = reference.index else {
            return None;
        };

        let Some(theme_elements) = self.theme_elements.clone() else {
            return None;
        };
        let Some(font_scheme) = theme_elements.font_scheme.clone() else {
            return None;
        };

        if index == "major" {
            return font_scheme.major_font;
        }

        if index == "minor" {
            return font_scheme.minor_font;
        };

        return None;
    }
}
