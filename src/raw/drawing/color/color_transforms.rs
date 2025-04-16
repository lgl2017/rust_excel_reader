use std::vec;

use crate::excel::XmlReader;
use crate::helper::{
    apply_modulation, apply_offset, apply_tint, complementary, extract_val_attribute, gamma_shift,
    grayscale, hsla_to_rgba, inverse, inverse_gamma_shift, percentage_int_to_float, rgba_to_hsla,
    string_to_bool, string_to_int,
};
use anyhow::bail;
use quick_xml::events::Event;

/// Common children configuration for drawingML colors
/// Example:
/// ```
/// <a:schemeClr val="phClr">
///     <a:alpha val="63000" />
///     <a:lumMod val="110000" />
///     <a:tint val="40000" />
///     <a:shade val="100000" />
///     <a:satMod val="350000" />
///     <a:comp/>
///     <a:inv/>
/// </a:schemeClr>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub enum ColorTransform {
    /// Alpha: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alpha?view=openxml-3.0.1
    // tag: alpha, attribute: val
    Alpha(i64),

    /// AlphaModulation: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphamodulation?view=openxml-3.0.1
    /// This element specifies a more or less opaque version of its input color.
    /// An alpha modulate never increases the alpha beyond 100%.
    /// A 200% alpha modulate makes a input color twice as opaque as before.
    /// A 50% alpha modulate makes a input color half as opaque as before.
    // tag: alphaMod, attribute: val
    AlphaModulation(i64),

    /// AlphaOffset: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoffset?view=openxml-3.0.1
    // tag: alphaOff, attribute: val
    AlphaOffset(i64),

    /// Blue: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blue?view=openxml-3.0.1
    // attribute: val
    Blue(i64),

    /// Blue Modification
    // tag: blueMod, attribute: val
    BlueModulation(i64),

    /// Blue Offset
    // tag: blueOff, attribute: val
    BlueOffset(i64),

    /// Green
    // attribute: val
    Green(i64),

    /// Green Modification
    // tag: greenMod, attribute: val
    GreenModulation(i64),
    /// Green Offset
    // tag: greenOff, attribute: val
    GreenOffset(i64),

    /// Red
    // tag: red,  attribute: val
    Red(i64),

    /// Red Modulation
    // tag: redMod,  attribute: val
    RedModulation(i64),

    /// Red Offset
    // tag: redOff,  attribute: val
    RedOffset(i64),

    // hue (Hue)	§20.1.2.3.14
    Hue(i64),

    // hueMod (Hue Modulate)	§20.1.2.3.15
    HueModulation(i64),

    // hueOff (Hue Offset)	§20.1.2.3.16
    HueOffset(i64),

    // lum (Luminance)	§20.1.2.3.19
    Lum(i64),

    // lumMod (Luminance Modulation)	§20.1.2.3.20
    LumModulation(i64),
    // lumOff (Luminance Offset)	§20.1.2.3.21
    LumOffset(i64),

    // sat (Saturation)	§20.1.2.3.26
    Sat(i64),
    // satMod (Saturation Modulation)	§20.1.2.3.27
    SatModulation(i64),
    // satOff (Saturation Offset)	§20.1.2.3.28
    SatOffset(i64),

    /// Complement: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.complement?view=openxml-3.0.1
    /// When presented, the color rendered should be the complement of its input color
    /// Example:
    /// ```
    /// <a:srgbClr val="FF0000">
    ///     <a:comp/>
    /// </a:srgbClr>
    /// ```
    // tag: comp
    Comp,

    /// Gamma: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gamma?view=openxml-3.0.1
    /// This element specifies that the output color rendered by the generating application should be the sRGB gamma shift of the input color.
    // tag: gamma, when presented, true
    Gamma,

    /// Gray: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gray?view=openxml-3.0.1
    /// This element specifies a grayscale of its input color, taking into relative intensities of the red, green, and blue primaries.
    // tag: gray, when presented, true
    Gray,

    /// Inverse: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.inverse?view=openxml-3.0.1
    /// This element specifies the inverse of its input color.
    /// Example:
    /// ```
    /// <a:srgbClr val="FF0000">
    ///     <a:inv/>
    ///  </a:srgbClr>
    /// ```
    // inv: gray, when presented, true
    Inverse,

    /// Inverse Gamma: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.inversegamma?view=openxml-3.0.1
    /// This element specifies that the output color rendered by the generating application should be the inverse sRGB gamma shift of the input color.
    // tag: invGamma, when presented, true
    InverseGamma,

    /// Shade: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shade?view=openxml-3.0.1
    /// specifies a darker version of its input color. A 10% shade is 10% of the input color combined with 90% black.
    /// Example:
    /// ```
    /// <a:srgbClr val="FF0000">
    ///     <a:shade val="100000" />
    ///  </a:srgbClr>
    /// ```
    // tag: shade, attribute: val
    Shade(i64),

    /// Tint: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tint?view=openxml-3.0.1
    /// This element specifies a lighter version of its input color. A 10% tint is 10% of the input color combined with 90% white.
    /// Example:
    /// ```
    /// <a:srgbClr val="FF0000">
    ///     <a:tint val="40000" />
    ///  </a:srgbClr>
    /// ```
    // tag: tint, attribute: val
    Tint(i64),
}

impl ColorTransform {
    pub fn load_list(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Vec<Self>> {
        let mut transforms: Vec<Self> = vec![];

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alpha" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Alpha(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::AlphaModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alphaOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::AlphaOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blue" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Blue(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blueMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::BlueModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blueOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::BlueOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"comp" => {
                    // <a:comp/>
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    if let Some(bool) = string_to_bool(&val_string) {
                        if bool {
                            transforms.push(ColorTransform::Comp);
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gamma" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    if let Some(bool) = string_to_bool(&val_string) {
                        if bool {
                            transforms.push(ColorTransform::Gamma);
                        }
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gray" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    if let Some(bool) = string_to_bool(&val_string) {
                        if bool {
                            transforms.push(ColorTransform::Gray);
                        }
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"green" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Green(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"greenMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::GreenModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"greenOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::GreenOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hue" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Hue(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hueMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::HueModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hueOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::HueOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"inv" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    if let Some(bool) = string_to_bool(&val_string) {
                        if bool {
                            transforms.push(ColorTransform::Inverse);
                        }
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"invGamma" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    if let Some(bool) = string_to_bool(&val_string) {
                        if bool {
                            transforms.push(ColorTransform::InverseGamma);
                        }
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lum" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Lum(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lumMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::LumModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lumOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::LumOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"red" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Red(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"redMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::RedModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"redOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::RedOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sat" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Sat(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"satMod" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::SatModulation(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"satOff" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::SatOffset(num));
                    }
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"shade" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Shade(num));
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tint" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    if let Some(num) = string_to_int(&val_string) {
                        transforms.push(ColorTransform::Tint(num));
                    }
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(transforms)
    }
}

impl ColorTransform {
    pub(crate) fn apply(&self, rgba: (u32, u32, u32, f64)) -> (u32, u32, u32, f64) {
        // let (mut r, mut g, mut b, mut a) = rgba;
        let (mut r, mut g, mut b) = (
            (rgba.0 as f64) / 255.0,
            (rgba.1 as f64) / 255.0,
            (rgba.2 as f64) / 255.0,
        );
        let mut a = rgba.3;

        match self {
            ColorTransform::Alpha(alpha) => a = percentage_int_to_float(*alpha),
            ColorTransform::AlphaModulation(modulation) => a = apply_modulation(a, *modulation),
            ColorTransform::AlphaOffset(offset) => a = apply_offset(a, *offset),

            ColorTransform::Blue(blue) => b = percentage_int_to_float(*blue),
            ColorTransform::BlueModulation(modulation) => b = apply_modulation(b, *modulation),
            ColorTransform::BlueOffset(offset) => b = apply_offset(b, *offset),

            ColorTransform::Green(green) => g = percentage_int_to_float(*green),
            ColorTransform::GreenModulation(modulation) => g = apply_modulation(g, *modulation),
            ColorTransform::GreenOffset(offset) => g = apply_offset(g, *offset),

            ColorTransform::Red(red) => r = percentage_int_to_float(*red),
            ColorTransform::RedModulation(modulation) => r = apply_modulation(r, *modulation),
            ColorTransform::RedOffset(offset) => r = apply_offset(r, *offset),

            ColorTransform::Hue(hue) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.0 = percentage_int_to_float(*hue) * 360.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::HueModulation(modulation) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.0 = apply_modulation(hsla.0 / 360.0, *modulation) * 360.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::HueOffset(offset) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.0 = apply_offset(hsla.0 / 360.0, *offset) * 360.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::Sat(sat) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.1 = percentage_int_to_float(*sat) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::SatModulation(modulation) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.1 = apply_modulation(hsla.1 / 100.0, *modulation) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::SatOffset(offset) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.1 = apply_offset(hsla.1 / 100.0, *offset) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }

            ColorTransform::Lum(lum) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.2 = percentage_int_to_float(*lum) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::LumModulation(modulation) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.2 = apply_modulation(hsla.2 / 100.0, *modulation) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::LumOffset(offset) => {
                if let Ok(mut hsla) = rgba_to_hsla(rgba) {
                    hsla.2 = apply_offset(hsla.2 / 100.0, *offset) * 100.0;
                    match hsla_to_rgba(hsla) {
                        Ok(new) => return new,
                        Err(_) => return rgba,
                    }
                }
            }
            ColorTransform::Comp => match complementary(rgba) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },
            ColorTransform::Gamma => match gamma_shift(rgba) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },

            ColorTransform::Gray => match grayscale(rgba) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },

            ColorTransform::Inverse => match inverse(rgba) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },
            ColorTransform::InverseGamma => match inverse_gamma_shift(rgba) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },

            ColorTransform::Shade(shade) => {
                match apply_tint(rgba, -percentage_int_to_float(*shade)) {
                    Ok(new) => return new,
                    Err(_) => return rgba,
                }
            }
            ColorTransform::Tint(tint) => match apply_tint(rgba, percentage_int_to_float(*tint)) {
                Ok(new) => return new,
                Err(_) => return rgba,
            },
        }

        return (
            (r * 255.0).round() as u32,
            (g * 255.0).round() as u32,
            (b * 255.0).round() as u32,
            a,
        );
    }
}

pub(crate) fn apply_color_transformations(
    rgba: (u32, u32, u32, f64),
    transformations: Vec<ColorTransform>,
) -> (u32, u32, u32, f64) {
    let mut rgba = rgba;
    for transform in transformations {
        rgba = transform.apply(rgba);
    }
    return rgba;
}
