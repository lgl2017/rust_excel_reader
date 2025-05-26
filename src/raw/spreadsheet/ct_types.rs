/// MAX_DIGIT_WIDTH Appximate with Aptos of font size 12 pt
static MAX_DIGIT_WIDTH: f64 = 8.2; // Max digit width: 7 pixel

/// Represent the following types:
/// * [CT_Font](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sz_topic_ID0E6DU5.html)
///
/// This element represents the point size (1/72 of an inch) of the Latin and East Asian text.
pub type CTFontSize = f64;

/// To conver the following two types to points.
/// - Column width: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_col_topic_ID0ELFQ4.html
/// - <defaultColWidth> (Default Column Width): https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html
///
/// This type Specifies the number of characters of the maximum digit width of the normal style's font.
/// Specifically, Column width measured as the number of characters of the maximum digit width of the numbers 0, 1, 2, ..., 9 as rendered in the normal style's font.
/// There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
///
/// width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
///
/// Using the Calibri font as an example, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi).
/// In fact, each digit is the same width for this font.
/// Therefore if the cell width is 8 characters wide, the value of this attribute shall be Truncate([8*7+5]/7*256)/256 = 8.7109375.
///
/// To translate the value of width in the file into the column width value at runtime (expressed in terms of pixels):
///
/// =Truncate((
/// (256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
///
/// Appximate with [MAX_DIGIT_WIDTH]: Aptos of font size 12 pt. -> Max digit width: 8.2 pixel
pub(crate) fn column_width_to_pt(w: f64) -> f64 {
    let px = (((256.0_f64 * w) + (128.0_f64 / MAX_DIGIT_WIDTH).trunc()) / 256.0_f64
        * MAX_DIGIT_WIDTH)
        .trunc();
    return px_to_pt(px);
}

/// To conver the following two types to points.
/// - <baseColWidth> (Base Column Width): https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sheetFormatPr_topic_ID0EVAG5.html
///
/// Specifies the number of characters of the maximum digit width of the normal style's font.
/// This value does not include margin padding or extra padding for gridlines.
/// It is only the number of characters.
///
/// To convert base column width to default column width:
/// width = Truncate([{@baseColumnWidth} * {Maximum Digit Width} + {5 pixel padding (2 pixels on each side, totalling 4 pixels + gridline (1pixel))}]/{Maximum Digit Width}*256)/256
///
/// Appximate with [MAX_DIGIT_WIDTH]: Aptos of font size 12 pt. -> Max digit width: 8.2 pixel
pub(crate) fn base_column_width_to_pt(base_column_width: u64) -> f64 {
    let base_column_width = base_column_width as f64;
    let default = ((base_column_width * MAX_DIGIT_WIDTH + 5.0_f64) / MAX_DIGIT_WIDTH * 256.0_f64)
        .trunc()
        / 256.0_f64;
    let px = column_width_to_pt(default);
    return px_to_pt(px);
}

fn px_to_pt(px: f64) -> f64 {
    return px * 72.0 / 96.0;
}
