use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colormap?view=openxml-3.0.1
///
/// Example:
/// ```
/// // bg1 is mapped to lt1, tx1 is mapped to dk1
/// <clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"  accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"  accent6="accent6" hlink="hlink" folHlink="folHlink"/>
/// ```
/// Possible values for all attributes:
/// ColorSchemeIndexValues: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorschemeindexvalues?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct ColorMap {
    /// attributes

    /// accent1 (Accent 1)	Specifies a color defined which is associated as the accent 1 color.
    pub accent1: Option<String>,

    /// accent2 (Accent 2)	Specifies a color defined which is associated as the accent 2 color.
    pub accent2: Option<String>,

    /// accent3 (Accent 3)	Specifies a color defined which is associated as the accent 3 color.
    pub accent3: Option<String>,

    /// accent4 (Accent 4)	Specifies a color defined which is associated as the accent 4 color.
    pub accent4: Option<String>,

    /// accent5 (Accent 5)	Specifies a color defined which is associated as the accent 5 color.
    pub accent5: Option<String>,

    /// accent6 (Accent 6)	Specifies a color defined which is associated as the accent 6 color.
    pub accent6: Option<String>,

    /// bg1 (Background 1)	A color defined which is associated as the first background color.
    pub bg1: Option<String>,

    /// bg2 (Background 2)	Specifies a color defined which is associated as the second background color.
    pub bg2: Option<String>,

    /// folHlink (Followed Hyperlink)	Specifies a color defined which is associated as the color for a followed hyperlink.
    pub fol_hlink: Option<String>,

    /// hlink (Hyperlink)	Specifies a color defined which is associated as the color for a hyperlink.
    pub hlink: Option<String>,

    /// tx1 (Text 1)	Specifies a color defined which is associated as the first text color.
    pub tx1: Option<String>,

    /// tx2 (Text 2)	Specifies a color defined which is associated as the second text color.
    pub tx2: Option<String>,
}

impl ColorMap {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut map = Self {
            accent1: None,
            accent2: None,
            accent3: None,
            accent4: None,
            accent5: None,
            accent6: None,
            bg1: None,
            bg2: None,
            fol_hlink: None,
            hlink: None,
            tx1: None,
            tx2: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"accent1" => map.accent1 = Some(string_value),
                        b"accent2" => map.accent2 = Some(string_value),
                        b"accent3" => map.accent3 = Some(string_value),
                        b"accent4" => map.accent4 = Some(string_value),
                        b"accent5" => map.accent5 = Some(string_value),
                        b"accent6" => map.accent6 = Some(string_value),
                        b"bg1" => map.bg1 = Some(string_value),
                        b"bg2" => map.bg2 = Some(string_value),
                        b"folHlink" => map.fol_hlink = Some(string_value),
                        b"hlink" => map.hlink = Some(string_value),
                        b"tx1" => map.tx1 = Some(string_value),
                        b"tx2" => map.tx2 = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(map)
    }
}
