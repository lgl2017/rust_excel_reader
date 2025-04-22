use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::effect::{
    alpha_bi_level::XlsxAlphaBiLevel, alpha_ceiling::XlsxAlphaCeiling, alpha_floor::XlsxAlphaFloor,
    alpha_inverse::XlsxAlphaInverse, alpha_modulation::XlsxAlphaModulation,
    alpha_modulation_fixed::XlsxAlphaModulationFixed, alpha_replace::XlsxAlphaReplace,
    bi_level::XlsxBiLevel, blur::XlsxBlur, color_change::XlsxColorChange,
    color_replacement::XlsxColorReplacement, duotone::XlsxDuotone, fill_overlay::XlsxFillOverlay,
    gray_scale::XlsxGrayScale, hue_saturation_luminance::XlsxHsl, luminance::XlsxLuminance,
    tint::XlsxTint,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBlip {
    //     Child Elements	Subclause
    // alphaBiLevel (Alpha Bi-Level Effect)	§20.1.8.1
    pub alpha_bi_level: Option<XlsxAlphaBiLevel>,

    // alphaCeiling (Alpha Ceiling Effect)	§20.1.8.2
    pub alpha_ceiling: Option<XlsxAlphaCeiling>,

    // alphaFloor (Alpha Floor Effect)	§20.1.8.3
    pub alpha_floor: Option<XlsxAlphaFloor>,

    // alphaInv (Alpha Inverse Effect)	§20.1.8.4
    pub alpha_inv: Option<XlsxAlphaInverse>,

    // alphaMod (Alpha Modulate Effect)	§20.1.8.5
    pub alpha_mod: Option<XlsxAlphaModulation>,

    // alphaModFix (Alpha Modulate Fixed Effect)	§20.1.8.6
    pub alpha_mod_fix: Option<XlsxAlphaModulationFixed>,

    // alphaRepl (Alpha Replace Effect)	§20.1.8.8
    pub alpha_repl: Option<XlsxAlphaReplace>,

    // biLevel (Bi-Level (Black/White) Effect)	§20.1.8.11
    pub bi_level: Option<XlsxBiLevel>,

    // blur (Blur Effect)	§20.1.8.15
    pub blur: Option<XlsxBlur>,

    // clrChange (Color Change Effect)	§20.1.8.16
    pub clr_change: Option<XlsxColorChange>,
    // clrRepl (Solid Color Replacement)	§20.1.8.18
    pub clr_repl: Option<XlsxColorReplacement>,

    // duotone (Duotone Effect)	§20.1.8.23
    pub duotone: Option<XlsxDuotone>,
    // fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<XlsxFillOverlay>>,
    // grayscl (Gray Scale Effect)	§20.1.8.34
    pub grayscl: Option<XlsxGrayScale>,

    // hsl (Hue Saturation Luminance Effect)	§20.1.8.39
    pub hsl: Option<XlsxHsl>,
    // lum (Luminance Effect)	§20.1.8.42
    pub lum: Option<XlsxLuminance>,
    // tint (Tint Effect)
    pub tint: Option<XlsxTint>,
    // extLst (Extension List)	§20.1.2.2.15 Not Supported

    // Attributes
    /// Specifies the compression state with which the picture is stored. This allows the application to specify the amount of compression that has been applied to a picture.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip?view=openxml-3.0.1
    // cstate (Compression State)
    pub cstate: Option<String>,

    /// Specifies the identification information for an embedded picture.
    /// This attribute is used to specify an image that resides locally within the file.
    ///
    /// Example:
    /// ```
    /// <a:blip r:embed="rId2"/>
    /// ```
    // embed (Embedded Picture Reference)
    pub embed: Option<String>,

    /// Specifies the identification information for a linked picture.
    /// This attribute is used to specify an image that does not reside within this file.
    // link (Linked Picture Reference)
    pub link: Option<String>,
}

impl XlsxBlip {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut blip = Self {
            alpha_bi_level: None,
            alpha_ceiling: None,
            alpha_floor: None,
            alpha_inv: None,
            alpha_mod: None,
            alpha_mod_fix: None,
            alpha_repl: None,
            bi_level: None,
            blur: None,
            clr_change: None,
            clr_repl: None,
            duotone: None,
            fill_overlay: None,
            grayscl: None,
            hsl: None,
            lum: None,
            tint: None,
            cstate: None,
            embed: None,
            link: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"cstate" => {
                            blip.cstate = Some(string_value);
                        }
                        b"embed" => {
                            blip.embed = Some(string_value);
                        }
                        b"link" => {
                            blip.link = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaBiLevel" => {
                    blip.alpha_bi_level = Some(XlsxAlphaBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaCeiling" => {
                    blip.alpha_ceiling = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaFloor" => {
                    blip.alpha_floor = Some(true);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaInv" => {
                    blip.alpha_inv = XlsxAlphaInverse::load(reader, b"alphaInv")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaMod" => {
                    blip.alpha_mod = Some(XlsxAlphaModulation::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaModFix" => {
                    blip.alpha_mod_fix = Some(XlsxAlphaModulationFixed::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaRepl" => {
                    blip.alpha_repl = Some(XlsxAlphaReplace::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"biLevel" => {
                    blip.bi_level = Some(XlsxBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blur" => {
                    blip.blur = Some(XlsxBlur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrChange" => {
                    blip.clr_change = Some(XlsxColorChange::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrRepl" => {
                    blip.clr_repl = XlsxColorReplacement::load(reader, b"clrRepl")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"duotone" => {
                    blip.duotone = Some(XlsxDuotone::load(reader)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = XlsxFillOverlay::load(reader, b"fillOverlay")? {
                        blip.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grayscl" => {
                    blip.grayscl = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hsl" => {
                    blip.hsl = Some(XlsxHsl::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lum" => {
                    blip.lum = Some(XlsxLuminance::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tint" => {
                    blip.tint = Some(XlsxTint::load(e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"blip" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(blip)
    }
}
