use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::effect::{
    alpha_bi_level::AlphaBiLevel, alpha_ceiling::AlphaCeiling, alpha_floor::AlphaFloor,
    alpha_inverse::AlphaInverse, alpha_modulation::AlphaModulation,
    alpha_modulation_fixed::AlphaModulationFixed, alpha_replace::AlphaReplace, bi_level::BiLevel,
    blur::Blur, color_change::ColorChange, color_replacement::ColorReplacement, duotone::Duotone,
    fill_overlay::FillOverlay, gray_scale::GrayScale, hue_saturation_luminance::Hsl,
    luminance::Luminance, tint::Tint,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct Blip {
    //     Child Elements	Subclause
    // alphaBiLevel (Alpha Bi-Level Effect)	§20.1.8.1
    pub alpha_bi_level: Option<AlphaBiLevel>,

    // alphaCeiling (Alpha Ceiling Effect)	§20.1.8.2
    pub alpha_ceiling: Option<AlphaCeiling>,

    // alphaFloor (Alpha Floor Effect)	§20.1.8.3
    pub alpha_floor: Option<AlphaFloor>,

    // alphaInv (Alpha Inverse Effect)	§20.1.8.4
    pub alpha_inv: Option<AlphaInverse>,

    // alphaMod (Alpha Modulate Effect)	§20.1.8.5
    pub alpha_mod: Option<AlphaModulation>,

    // alphaModFix (Alpha Modulate Fixed Effect)	§20.1.8.6
    pub alpha_mod_fix: Option<AlphaModulationFixed>,

    // alphaRepl (Alpha Replace Effect)	§20.1.8.8
    pub alpha_repl: Option<AlphaReplace>,

    // biLevel (Bi-Level (Black/White) Effect)	§20.1.8.11
    pub bi_level: Option<BiLevel>,

    // blur (Blur Effect)	§20.1.8.15
    pub blur: Option<Blur>,

    // clrChange (Color Change Effect)	§20.1.8.16
    pub clr_change: Option<ColorChange>,
    // clrRepl (Solid Color Replacement)	§20.1.8.18
    pub clr_repl: Option<ColorReplacement>,

    // duotone (Duotone Effect)	§20.1.8.23
    pub duotone: Option<Duotone>,
    // fillOverlay (Fill Overlay Effect)	§20.1.8.29
    pub fill_overlay: Option<Box<FillOverlay>>,
    // grayscl (Gray Scale Effect)	§20.1.8.34
    pub grayscl: Option<GrayScale>,

    // hsl (Hue Saturation Luminance Effect)	§20.1.8.39
    pub hsl: Option<Hsl>,
    // lum (Luminance Effect)	§20.1.8.42
    pub lum: Option<Luminance>,
    // tint (Tint Effect)
    pub tint: Option<Tint>,
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

impl Blip {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
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
                    blip.alpha_bi_level = Some(AlphaBiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaCeiling" => {
                    blip.alpha_ceiling = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaFloor" => {
                    blip.alpha_floor = Some(true);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaInv" => {
                    blip.alpha_inv = AlphaInverse::load(reader, b"alphaInv")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaMod" => {
                    blip.alpha_mod = Some(AlphaModulation::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaModFix" => {
                    blip.alpha_mod_fix = Some(AlphaModulationFixed::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaRepl" => {
                    blip.alpha_repl = Some(AlphaReplace::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"biLevel" => {
                    blip.bi_level = Some(BiLevel::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blur" => {
                    blip.blur = Some(Blur::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrChange" => {
                    blip.clr_change = Some(ColorChange::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrRepl" => {
                    blip.clr_repl = ColorReplacement::load(reader, b"clrRepl")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"duotone" => {
                    blip.duotone = Some(Duotone::load(reader)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillOverlay" => {
                    if let Some(fill_overlay) = FillOverlay::load(reader, b"fillOverlay")? {
                        blip.fill_overlay = Some(Box::new(fill_overlay));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grayscl" => {
                    blip.grayscl = Some(true);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hsl" => {
                    blip.hsl = Some(Hsl::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lum" => {
                    blip.lum = Some(Luminance::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tint" => {
                    blip.tint = Some(Tint::load(e)?);
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
