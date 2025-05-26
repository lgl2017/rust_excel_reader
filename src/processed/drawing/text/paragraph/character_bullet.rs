#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::text::paragraph::character_bullet::XlsxCharacterBullet;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.characterbullet?view=openxml-3.0.1
///
/// This element specifies that a character be applied to a set of bullets.
/// These bullets are allowed to be any character in any font that the system is able to support.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buFont typeface="Calibri"/>
///     <a:buChar char="g"/>
/// </a:pPr>
/// ```
// tag: buChar
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct CharacterBullet {
    ///	Specifies the character to be used in place of the standard bullet point.
    /// This character can be any character for the specified font that is supported by the system upon which this document is being viewed.
    pub character: String,
}

impl CharacterBullet {
    pub(crate) fn from_raw(raw: XlsxCharacterBullet) -> Option<Self> {
        let Some(c) = raw.char else { return None };

        return Some(Self { character: c });
    }
}
