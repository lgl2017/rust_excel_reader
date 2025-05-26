use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguide?view=openxml-3.0.1
///
/// This element specifies the precense of a shape guide that is used to govern the geometry of the specified shape.
/// A shape guide consists of a formula and a name that the result of the formula is assigned to.
/// Recognized formulas are listed with the fmla attribute documentation for this element.
///
/// Example:
/// ```
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst>
///     <a:gd name="myGuide" fmla="*/ h 2 3"/>
///   </a:gdLst>
///   <a:ahLst/>
///   <a:cxnLst/>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="1705233" h="679622">
///       <a:moveTo>
///         <a:pt x="0" y="myGuide"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="1705233" y="myGuide"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="852616" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
///
/// gd (Shpae guide)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShapeGuide {
    // Attributes
    /// Specifies the formula that is used to calculate the value for a guide.
    /// Each formula has a certain number of arguments and a specific set of operations to perform on these arguments in order to generate a value for a guide.
    /// formulas available: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguide?view=openxml-3.0.1
    // fmla (Shape Guide Formula)
    pub formula: Option<String>,

    /// Specifies the name that is used to reference to this guide.
    /// This name can be used just as a variable would within an equation.
    // name (Shape Guide Name)
    pub name: Option<String>,
}

impl XlsxShapeGuide {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut guide = Self {
            formula: None,
            name: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"fmla" => guide.formula = Some(string_value),
                        b"name" => guide.name = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(guide)
    }
}
