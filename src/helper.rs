use anyhow::bail;
use chrono::{DateTime, NaiveDateTime};
use quick_xml::events::BytesStart;
use regex::Regex;

use crate::common_types::XlsxDatetime;

/// Converting Attributes string to datetime
pub(crate) fn string_to_datetime(str: &str) -> Option<XlsxDatetime> {
    // with time zone: YYYY-MM-DDThh:mm:ssZ
    if let Ok(date_time) = DateTime::parse_from_rfc3339(str) {
        return Some(XlsxDatetime {
            datetime: date_time.naive_utc(),
            offset: Some(date_time.offset().to_owned()),
        });
    }

    // without time zone: YYYY-MM-DDThh:mm:ss
    if let Ok(naive_date_time) = NaiveDateTime::parse_from_str(&str, "%Y-%m-%dT%H:%M:%S") {
        return Some(XlsxDatetime {
            datetime: naive_date_time,
            offset: None,
        });
    };

    return None;
}

/// Converting Attributes string to boolean
pub(crate) fn string_to_bool(str: &str) -> Option<bool> {
    return match str {
        "0" | "false" => Some(false),
        "1" | "true" => Some(true),
        _ => None,
    };
}

/// Converting Attributes string to Unsigned Int
pub(crate) fn string_to_unsignedint(str: &str) -> Option<u64> {
    return match str.parse::<u64>() {
        Ok(int) => Some(int),
        Err(_) => None,
    };
}

/// Converting Attributes string to Integer
pub(crate) fn string_to_int(str: &str) -> Option<i64> {
    if str.ends_with("%") {
        return percentage_string_to_int(str);
    }
    return match str.parse::<i64>() {
        Ok(int) => Some(int),
        Err(_) => None,
    };
}

/// Converting percentage Integer to float.
///
/// Ex: 56,000 => 0.56
pub(crate) fn percentage_int_to_float(int: i64) -> f64 {
    return (int as f64) / 1000.0 / 100.0;
}

/// Converting Attributes percentage to integar.
///
/// 100.000% -> 100000
fn percentage_string_to_int(str: &str) -> Option<i64> {
    let mut chars = str.chars();
    chars.next_back();
    let str = chars.as_str();

    return match str.parse::<f64>() {
        Ok(f) => Some((f * 1000.0) as i64),
        Err(_) => None,
    };
}

/// Converting Attributes string to float
pub(crate) fn string_to_float(str: &str) -> Option<f64> {
    return match str.parse::<f64>() {
        Ok(float) => Some(float),
        Err(_) => None,
    };
}

/// Extract `val` from attributes
pub(crate) fn extract_val_attribute(e: &BytesStart) -> anyhow::Result<Option<String>> {
    let attributes = e.attributes();
    for a in attributes {
        match a {
            Ok(a) => {
                let string_value = String::from_utf8(a.value.to_vec())?;
                match a.key.local_name().as_ref() {
                    b"val" => return Ok(Some(string_value)),
                    _ => {}
                }
            }
            Err(error) => {
                bail!(error.to_string())
            }
        }
    }

    return Ok(None);
}

/// Convert A1 reference dimension to (row, col) (1 based index).
///
/// - top left (row, column),
/// - bottom right (row, column)
pub(crate) fn a1_dimension_to_row_col(
    a1_dimension: &[u8],
) -> anyhow::Result<((u64, u64), (u64, u64))> {
    let parts: Vec<&[u8]> = a1_dimension.split(|c| *c == b':').collect();
    if parts.len() != 2 {
        bail!("Invalid reference dimension.")
    }
    let (Some(start_row), Some(start_col)) = a1_address_to_row_col(parts[0])? else {
        bail!("Invalid reference start address.")
    };
    let (Some(end_row), Some(end_col)) = a1_address_to_row_col(parts[1])? else {
        bail!("Invalid reference end address.")
    };

    Ok(((start_row, start_col), (end_row, end_col)))
}

/// Convert R1C1 reference dimension to (row, col) (1 based index).
///
/// - top left (row, column),
/// - bottom right (row, column)
pub(crate) fn r1c1_dimension_to_row_col(
    r1c1_dimension: &str,
) -> anyhow::Result<((u64, u64), (u64, u64))> {
    let parts: Vec<&str> = r1c1_dimension.split(|c: char| c == ':').collect();
    if parts.len() != 2 {
        bail!("Invalid reference dimension.")
    }
    let Some(start) = r1c1_address_to_row_col(parts[0])? else {
        bail!("Invalid reference start address.")
    };
    let Some(end) = r1c1_address_to_row_col(parts[1])? else {
        bail!("Invalid reference end address.")
    };

    Ok((start, end))
}

/// Convert R1C1 reference to (row, col) (1 based index).
pub(crate) fn r1c1_address_to_row_col(r1c1: &str) -> anyhow::Result<Option<(u64, u64)>> {
    let r1c1 = r1c1.to_ascii_uppercase();
    let re = Regex::new(r"R(?<row>[0-9]+)C(?<col>[0-9]+)$")?;
    let Some(caps) = re.captures(&r1c1) else {
        return Ok(None);
    };
    let row = &caps["row"].parse::<u64>()?;
    let col = &caps["col"].parse::<u64>()?;
    return Ok(Some((*row, *col)));
}

/// Convert A1 reference to (row, col) (1 based index).
/// A6 -> col: 1, row: 6
/// A -> col: 1, row: None (Row does not exist)
/// 6 -> col: None (col does not exist), row: 6
pub(crate) fn a1_address_to_row_col(
    a1_address: &[u8],
) -> anyhow::Result<(Option<u64>, Option<u64>)> {
    let (mut row, mut col) = (0, 0);
    let mut power = 1;
    let mut reading_row = true;
    for c in a1_address.iter().rev() {
        match *c {
            c @ b'0'..=b'9' => {
                if reading_row {
                    row += ((c - b'0') as u64) * power;
                    power *= 10;
                } else {
                    bail!("Cell address contains numeric column.")
                }
            }
            c @ b'A'..=b'Z' => {
                if reading_row {
                    power = 1;
                    reading_row = false;
                }
                col += ((c - b'A') as u64 + 1) * power;
                power *= 26;
            }
            c @ b'a'..=b'z' => {
                if reading_row {
                    power = 1;
                    reading_row = false;
                }
                col += ((c - b'a') as u64 + 1) * power;
                power *= 26;
            }
            _ => bail!("Cell address is not alphaNumeric."),
        }
    }
    let row = if row.eq(&0) { None } else { Some(row) };
    let col = if col.eq(&0) { None } else { Some(col) };

    Ok((row, col))
}

/// Format hex string to RGBA hex string, ie: #960d52ff
pub(crate) fn format_hex_string(hex: &str, alpha_first: Option<bool>) -> anyhow::Result<String> {
    let mut s = hex;
    if s.starts_with("#") {
        s = s.trim_start_matches("#");
    }
    if s.len() != 6 && s.len() != 8 {
        bail!("invalid hex.")
    }
    let alpha_first = alpha_first.unwrap_or(false);
    if !alpha_first && s.len() == 8 {
        return Ok(format!("#{}", hex.to_ascii_lowercase()));
    }

    if s.len() == 6 {
        return Ok(format!("#{}{}", hex.to_ascii_lowercase(), "ff"));
    }
    let alpha_hex = &hex[..=1];
    let rgb_hex = &hex[2..];

    return Ok(format!("#{}{}", rgb_hex, alpha_hex));
}

/// hex string to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub(crate) fn hex_to_rgba(
    hex_str: &str,
    alpha_first: Option<bool>,
) -> anyhow::Result<(u32, u32, u32, f64)> {
    let mut s = hex_str;
    if s.starts_with("#") {
        s = s.trim_start_matches("#");
    }
    if s.len() != 6 && s.len() != 8 {
        bail!("invalid hex.")
    }

    // hex without alpha
    if s.len() == 6 {
        let num = u32::from_str_radix(s, 16)?;
        let r = (num & 0xFF0000) >> 16;
        let g = (num & 0x00FF00) >> 8;
        let b = (num & 0x0000FF) >> 0;
        return Ok((r, g, b, 1.0));
    }

    // alpha first
    if alpha_first.unwrap_or(false) {
        let num = u32::from_str_radix(s, 16)?;
        let a = (num & 0xFF000000) >> 24;
        let r = (num & 0x00FF0000) >> 16;
        let g = (num & 0x0000FF00) >> 8;
        let b = (num & 0x000000FF) >> 0;
        return Ok((r, g, b, (a as f64) / 255.0));
    }

    // alpha last
    let num = u32::from_str_radix(s, 16)?;
    let r = (num & 0xFF000000) >> 24;
    let g = (num & 0x00FF0000) >> 16;
    let b = (num & 0x0000FF00) >> 8;
    let a = (num & 0x000000FF) >> 0;
    return Ok((r, g, b, (a as f64) / 255.0));
}

/// RGBA value to hex string
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub(crate) fn rgba_to_hex(
    rgba: (u32, u32, u32, f64),
    alpha_first: Option<bool>,
) -> anyhow::Result<String> {
    let (r, g, b, a) = rgba;
    if !check_rgb(&r) || !check_rgb(&g) || !check_rgb(&b) {
        bail!("invalid rgb value.")
    }

    fn num_to_hex(n: u32) -> String {
        let s = format!("{:x}", n);
        if s.len() == 1 {
            String::from("0") + &s
        } else {
            s
        }
    }

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }

    let a_u32 = (a * 255.0).round() as u32;

    // alpha first
    if alpha_first.unwrap_or(false) {
        return Ok(format!(
            "#{}{}{}{}",
            num_to_hex(a_u32),
            num_to_hex(r),
            num_to_hex(g),
            num_to_hex(b)
        ));
    }

    // alpha last
    return Ok(format!(
        "#{}{}{}{}",
        num_to_hex(r),
        num_to_hex(g),
        num_to_hex(b),
        num_to_hex(a_u32)
    ));
}

/// HSLA to RGBA
///
/// Ranges:
/// * H: 0.0 - 360.0
/// * S: 0.0 - 100.0
/// * L: 0.0 - 100.0
///
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
///
/// * A: 0.0 - 1.0
pub(crate) fn hsla_to_rgba(hsla: (f64, f64, f64, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    if !check_hue(&hsla.0) || !check_percentage(&hsla.1) || !check_percentage(&hsla.2) {
        bail!("invalid hsl value.")
    }
    let (h, s, l, a) = (hsla.0 / 360.0, hsla.1 / 100.0, hsla.2 / 100.0, hsla.3);

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }
    if s == 0.0 {
        return Ok((255, 255, 255, a));
    }

    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };

    let p = 2.0 * l - q;

    fn calc(p: f64, q: f64, t: f64) -> f64 {
        let mut t = t;
        if t < 0.0 {
            t += 1.0
        };
        if t > 1.0 {
            t -= 1.0
        };
        if t < 1.0 / 6.0 {
            return p + (q - p) * 6.0 * t;
        };
        if t < 1.0 / 2.0 {
            return q;
        };
        if t < 2.0 / 3.0 {
            return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
        };
        return p;
    }
    let r = calc(p, q, h + 1.0 / 3.0);
    let g = calc(p, q, h);
    let b = calc(p, q, h - 1.0 / 3.0);

    return Ok((
        (r * 255.0).round() as u32,
        (g * 255.0).round() as u32,
        (b * 255.0).round() as u32,
        a,
    ));
}

/// RGBA to HSLA
///
/// Ranges:
/// * H: 0.0 - 360.0
/// * S: 0.0 - 100.0
/// * L: 0.0 - 100.0
///
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
///
/// * A: 0.0 - 1.0
pub(crate) fn rgba_to_hsla(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(f64, f64, f64, f64)> {
    if !check_rgb(&rgba.0) || !check_rgb(&rgba.1) || !check_rgb(&rgba.2) {
        bail!("invalid rgb value.")
    }
    let (r, g, b, a) = (
        (rgba.0 as f64) / 255.0,
        (rgba.1 as f64) / 255.0,
        (rgba.2 as f64) / 255.0,
        rgba.3,
    );

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }

    let (max, max_index) = max(vec![r, g, b]).unwrap_or((1.0, 0));
    let (min, _) = min(vec![r, g, b]).unwrap_or((0.0, 0));

    let lum = (max + min) / 2.0;
    if max == min {
        return Ok((0.0, 0.0, lum, a));
    }

    let chroma = max - min;
    let sat = chroma / (1.0 - (2.0 * lum - 1.0).abs());

    let hue = match max_index {
        // r
        0 => {
            let x = if g < b { 6.0 } else { 0.0 };
            (g - b) / chroma + x
        }
        // g
        1 => (b - r) / chroma + 2.0,
        // b
        2 => (r - g) / chroma + 4.0,
        _ => unreachable!(),
    };

    let mut hue = hue * 60.0;
    while hue < 0.0 {
        hue = hue + 360.0
    }
    while hue >= 360.0 {
        hue = hue - 360.0
    }

    return Ok((hue, sat * 100.0, lum * 100.0, a));
}

/// HSVA to RGBA
///
/// Ranges:
/// * H: 0.0 - 360.0 (exclusive)
/// * S: 0.0 - 100.0
/// * V: 0.0 - 100.0
///
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
///
/// * A: 0.0 - 1.0
pub fn hsva_to_rgba(hsva: (f64, f64, f64, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    if !check_hue(&hsva.0) || !check_percentage(&hsva.1) || !check_percentage(&hsva.2) {
        bail!("invalid hsv value.")
    }
    let (h, s, v, a) = (hsva.0 / 360.0, hsva.1 / 100.0, hsva.2 / 100.0, hsva.3);

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }

    let i = (h * 6.0).floor();
    let f = h * 6.0 - i;
    let p = v * (1.0 - s);
    let q = v * (1.0 - f * s);
    let t = v * (1.0 - (1.0 - f) * s);

    if s == 0.0 {
        return Ok((255, 255, 255, a));
    }

    let r: f64;
    let g: f64;
    let b: f64;

    match (i as u32) % 6 {
        0 => {
            r = v;
            g = t;
            b = p;
        }
        1 => {
            r = q;
            g = v;
            b = p;
        }
        2 => {
            r = p;
            g = v;
            b = t;
        }
        3 => {
            r = p;
            g = q;
            b = v;
        }
        4 => {
            r = t;
            g = p;
            b = v;
        }
        5 => {
            r = v;
            g = p;
            b = q;
        }
        _ => unreachable!(),
    }

    return Ok((
        (r * 255.0).round() as u32,
        (g * 255.0).round() as u32,
        (b * 255.0).round() as u32,
        a,
    ));
}

/// RGBA to HSVA
///
/// Ranges:
/// * H: 0.0 - 360.0 (exclusive)
/// * S: 0.0 - 100.0
/// * V: 0.0 - 100.0
///
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
///
/// * A: 0.0 - 1.0
pub fn rgba_to_hsva(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(f64, f64, f64, f64)> {
    if !check_rgb(&rgba.0) || !check_rgb(&rgba.1) || !check_rgb(&rgba.2) {
        bail!("invalid rgb value.")
    }
    let (r, g, b, a) = (
        (rgba.0 as f64) / 255.0,
        (rgba.1 as f64) / 255.0,
        (rgba.2 as f64) / 255.0,
        rgba.3,
    );

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }

    let (max, max_index) = max(vec![r, g, b]).unwrap_or((1.0, 0));
    let (min, _) = min(vec![r, g, b]).unwrap_or((0.0, 0));
    let val = max;

    if max == min {
        return Ok((0.0, 0.0, val, a));
    }

    let diff = max - min;
    let hue: f64;

    let sat = diff / val;

    match max_index {
        0 => hue = (g - b) / diff,
        1 => hue = 2.0 + (b - r) / diff,
        2 => hue = 4.0 + (r - g) / diff,
        _ => unreachable!(),
    }
    let mut hue = hue * 60.0;
    while hue < 0.0 {
        hue = hue + 360.0
    }
    while hue >= 360.0 {
        hue = hue - 360.0
    }

    return Ok((hue, sat * 100.0, val * 100.0, a));
}

/// Complementary Color
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub fn complementary(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    let (mut h, s, v, a) = rgba_to_hsva(rgba)?;
    fn shift_hue(h: f64, s: f64) -> f64 {
        let mut h = h + s;
        while h < 0.0 {
            h = h + 360.0
        }
        while h >= 360.0 {
            h = h - 360.0
        }
        return h;
    }
    h = shift_hue(h, 180.0);
    return hsva_to_rgba((h, s, v, a));
}

/// Inverse Color
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub fn inverse(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    if !check_rgb(&rgba.0) || !check_rgb(&rgba.1) || !check_rgb(&rgba.2) {
        bail!("invalid rgb value.")
    }

    if !check_alpha(&rgba.3) {
        bail!("invalid alpha value")
    }

    return Ok((255 - rgba.0, 255 - rgba.0, 255 - rgba.0, rgba.3));
}

/// SRGB Gamma shifted color
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub(crate) fn gamma_shift(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    let (h, s, mut v, a) = rgba_to_hsva(rgba)?;
    v = v.powf(2.2);
    return hsva_to_rgba((h, s, v, a));
}

/// SRGB inverse Gamma shifted color
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub(crate) fn inverse_gamma_shift(
    rgba: (u32, u32, u32, f64),
) -> anyhow::Result<(u32, u32, u32, f64)> {
    let (h, s, mut v, a) = rgba_to_hsva(rgba)?;
    v = v.powf(1.0 / 2.2);
    return hsva_to_rgba((h, s, v, a));
}

/// Get GrayScale of a color
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
pub(crate) fn grayscale(rgba: (u32, u32, u32, f64)) -> anyhow::Result<(u32, u32, u32, f64)> {
    if !check_rgb(&rgba.0) || !check_rgb(&rgba.1) || !check_rgb(&rgba.2) {
        bail!("invalid rgb value.")
    }
    let (r, g, b, a) = (
        (rgba.0 as f64) * 0.299,
        (rgba.1 as f64) * 0.587,
        (rgba.2 as f64) * 0.114,
        rgba.3,
    );

    if !check_alpha(&a) {
        bail!("invalid alpha value")
    }

    let gray = (r + g + b).round() as u32;

    Ok((gray, gray, gray, a))
}

/// Apply tint to a specific color.
///
/// RGBA to RGBA
///
/// Ranges:
/// * R: 0 - 255
/// * G: 0 - 255
/// * B: 0 - 255
/// * A: 0.0 - 1.0
///
/// tint: -1.0 to 1.0.
/// * when less than 0, specifies a darker version of the input color. -1.0 means 100% darken
/// * when greated than 0, specifies a lighter version of its input color. 1.0 means 100% lighten.
/// * 0.0 means no change.
pub(crate) fn apply_tint(
    rgba: (u32, u32, u32, f64),
    tint: f64,
) -> anyhow::Result<(u32, u32, u32, f64)> {
    if tint == 0.0 {
        return Ok((rgba.0, rgba.1, rgba.2, rgba.3));
    }
    let (h, s, mut l, a) = rgba_to_hsla(rgba)?;
    l = l / 100.0;

    if tint < 0.0 {
        l = l * (1.0 + tint);
    } else {
        l = l * (1.0 - tint) + (1.0 - (1.0 - tint));
    }

    if l < 0.0 {
        l = 0.0
    }
    if l > 1.0 {
        l = 1.0
    }

    return hsla_to_rgba((h, s, l, a));
}

/// apply modulation (percentage) to a specific color component
///
/// Example: blue modulation (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bluemodulation?view=openxml-3.0.1)
/// A 50% blue modulate reduces the blue component by half.
/// A 200% blue modulate doubles the blue component.
///
/// original/returned: color component ranging from 0.0 to 1.0
/// modulation: percentage in int. Ex: 50% -> 50,000
pub(crate) fn apply_modulation(original: f64, modulation: i64) -> f64 {
    let modulation = percentage_int_to_float(modulation);
    let mut new = original * modulation;
    if new < 0.0 {
        new = 0.0
    }
    if new > 1.0 {
        new = 1.0
    }
    return new;
}

/// apply offset (percetage) to a specific color component
///
/// Example: alpha offset (https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphaoffset?view=openxml-3.0.1)
/// A 10% alpha offset increases a 50% opacity to 60%.
/// A -10% alpha offset decreases a 50% opacity to 40%.
///
/// original/returned: color component ranging from 0.0 to 1.0
/// offset: percentage in int. Ex: 50% -> 50,000
pub(crate) fn apply_offset(original: f64, offset: i64) -> f64 {
    let mut new = original + percentage_int_to_float(offset);
    if new < 0.0 {
        new = 0.0
    }
    if new > 1.0 {
        new = 1.0
    }
    return new;
}

/********************************/
/****** Helper functions ********/
/********************************/
/// check rgb value
fn check_rgb(num: &u32) -> bool {
    return (0..=255).contains(num);
}

/// check alpha value
fn check_alpha(a: &f64) -> bool {
    return (0.0..=1.0).contains(a);
}

/// check hue value
fn check_hue(h: &f64) -> bool {
    return (0.0..=360.0).contains(h);
}

/// check percentage
fn check_percentage(n: &f64) -> bool {
    return (0.0..=100.0).contains(n);
}

/// min with index of a vector
fn min<T: std::cmp::PartialOrd + Copy>(v: Vec<T>) -> Option<(T, usize)> {
    if v.is_empty() {
        return None;
    }
    let mut result = v[0];
    let mut index: usize = 0;
    if v.len() == 1 {
        return Some((result, index));
    }
    for i in 1..v.len() {
        let v = v[i];
        if v < result {
            result = v;
            index = i;
        }
    }
    Some((result, index))
}

/// min with index of a vector
fn max<T: std::cmp::PartialOrd + Copy>(v: Vec<T>) -> Option<(T, usize)> {
    if v.is_empty() {
        return None;
    }
    let mut result = v[0];
    let mut index: usize = 0;
    if v.len() == 1 {
        return Some((result, index));
    }
    for i in 1..v.len() {
        let v = v[i];
        if v > result {
            result = v;
            index = i;
        }
    }
    Some((result, index))
}
