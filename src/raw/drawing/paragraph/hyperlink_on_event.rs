use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::helper::string_to_bool;

use super::hyperlink_sound::XlsxHyperlinkSound;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinkonmouseover?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:hlinkMouseOver r:id="rId2" tooltip="Some Sample Text"/>
/// ```
// tag: hlinkMouseOver
pub type XlsxHyperlinkOnMouseOver = XlsxHyperlinkOnEvent;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinkonclick?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:hlinkClick r:id="rId2" tooltip="Some Sample Text"/>
/// ```
// tag: hlinkClick
pub type XlsxHyperlinkOnClick = XlsxHyperlinkOnEvent;

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHyperlinkOnEvent {
    // extLst (Extension List)	Not supported

    //  Child Elements	Subclause
    // snd (Hyperlink Sound)
    pub sound: Option<XlsxHyperlinkSound>,

    // Attributes
    /// Specifies an action that is to be taken when this hyperlink is activated.
    /// This can be used to specify a slide to be navigated to or a script of code to be run.
    // action (Action Setting)
    pub action: Option<String>,

    /// Specifies if the URL in question should stop all sounds that are playing when it is clicked.
    // endSnd
    pub end_sound: Option<bool>,

    /// Specifies if this attribute has already been used within this document.
    // highlightClick
    pub highlight_click: Option<bool>,

    /// Specifies whether to add this to the history when navigating to it.
    // history
    pub history: Option<bool>,

    /// Specifies the relationship id that when looked up in this slides relationship file contains the target of this hyperlink.
    /// This attribute cannot be omitted.
    // id (Drawing Object Hyperlink Target)
    pub id: Option<String>,

    /// Specifies the URL when it has been determined by the generating application that the URL is invalid.
    // invalidUrl
    pub invalid_url: Option<String>,

    /// Specifies the target frame that is to be used when opening this hyperlink.
    /// When the hyperlink is activated this attribute is used to determine if a new window is launched for viewing or if an existing one can be used.
    /// If this attribute is omitted, than a new window is opened.
    // tgtFrame
    pub target_frame: Option<String>,

    /// Specifies the tooltip that should be displayed when the hyperlink text is hovered over with the mouse. If this attribute is omitted, than the hyperlink text itself can be displayed.
    // tooltip (Hyperlink Tooltip)
    pub tooltip: Option<String>,
}

impl XlsxHyperlinkOnEvent {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut onclick: Self = Self {
            sound: None,
            // attributes
            action: None,
            end_sound: None,
            highlight_click: None,
            history: None,
            id: None,
            invalid_url: None,
            target_frame: None,
            tooltip: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"action" => {
                            onclick.action = Some(string_value);
                        }
                        b"endSnd" => {
                            onclick.end_sound = string_to_bool(&string_value);
                        }
                        b"highlightClick" => {
                            onclick.highlight_click = string_to_bool(&string_value);
                        }
                        b"history" => {
                            onclick.history = string_to_bool(&string_value);
                        }
                        b"id" => {
                            onclick.id = Some(string_value);
                        }
                        b"invalidUrl" => {
                            onclick.invalid_url = Some(string_value);
                        }
                        b"tgtFrame" => {
                            onclick.target_frame = Some(string_value);
                        }
                        b"tooltip" => {
                            onclick.tooltip = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"snd" => {
                    onclick.sound = Some(XlsxHyperlinkSound::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"hlinkClick" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(onclick)
    }
}
