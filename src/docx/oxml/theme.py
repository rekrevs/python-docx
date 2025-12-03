"""Custom element classes for theme-related elements."""

from __future__ import annotations

from docx.oxml.simpletypes import XsdString
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrOne


class CT_OfficeStyleSheet(BaseOxmlElement):
    """``<a:theme>`` element, the root element for a theme part.

    Defines the color scheme, font scheme, and format scheme (effects) for a document.
    """

    themeElements: CT_BaseStyles | None = ZeroOrOne("a:themeElements")  # pyright: ignore[reportAssignmentType]
    name: str | None = OptionalAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_BaseStyles(BaseOxmlElement):
    """``<a:themeElements>`` element.

    Contains the color scheme, font scheme, and format scheme elements.
    """

    clrScheme: CT_ColorScheme | None = ZeroOrOne("a:clrScheme")  # pyright: ignore[reportAssignmentType]
    fontScheme: CT_FontScheme | None = ZeroOrOne("a:fontScheme")  # pyright: ignore[reportAssignmentType]


class CT_ColorScheme(BaseOxmlElement):
    """``<a:clrScheme>`` element.

    Defines the 12 theme colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
    """

    name: str | None = OptionalAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]

    # Theme color elements
    dk1: CT_Color | None = ZeroOrOne("a:dk1")  # pyright: ignore[reportAssignmentType]
    lt1: CT_Color | None = ZeroOrOne("a:lt1")  # pyright: ignore[reportAssignmentType]
    dk2: CT_Color | None = ZeroOrOne("a:dk2")  # pyright: ignore[reportAssignmentType]
    lt2: CT_Color | None = ZeroOrOne("a:lt2")  # pyright: ignore[reportAssignmentType]
    accent1: CT_Color | None = ZeroOrOne("a:accent1")  # pyright: ignore[reportAssignmentType]
    accent2: CT_Color | None = ZeroOrOne("a:accent2")  # pyright: ignore[reportAssignmentType]
    accent3: CT_Color | None = ZeroOrOne("a:accent3")  # pyright: ignore[reportAssignmentType]
    accent4: CT_Color | None = ZeroOrOne("a:accent4")  # pyright: ignore[reportAssignmentType]
    accent5: CT_Color | None = ZeroOrOne("a:accent5")  # pyright: ignore[reportAssignmentType]
    accent6: CT_Color | None = ZeroOrOne("a:accent6")  # pyright: ignore[reportAssignmentType]
    hlink: CT_Color | None = ZeroOrOne("a:hlink")  # pyright: ignore[reportAssignmentType]
    folHlink: CT_Color | None = ZeroOrOne("a:folHlink")  # pyright: ignore[reportAssignmentType]


class CT_Color(BaseOxmlElement):
    """Theme color element like ``<a:dk1>``, ``<a:accent1>``, etc.

    Contains either a system color (sysClr) or an RGB color (srgbClr).
    """

    sysClr: CT_SystemColor | None = ZeroOrOne("a:sysClr")  # pyright: ignore[reportAssignmentType]
    srgbClr: CT_SRgbColor | None = ZeroOrOne("a:srgbClr")  # pyright: ignore[reportAssignmentType]

    @property
    def rgb_color(self) -> str | None:
        """Return the RGB color value as a hex string (e.g., 'FF0000').

        For system colors, returns the lastClr attribute which represents the
        actual color value when the document was last opened.
        """
        if self.srgbClr is not None:
            return self.srgbClr.val
        if self.sysClr is not None:
            return self.sysClr.lastClr
        return None

    @rgb_color.setter
    def rgb_color(self, value: str) -> None:
        """Set the RGB color value, replacing any existing color definition.

        Args:
            value: Hex color string without '#' (e.g., 'FF0000' for red).
        """
        # Remove any existing color children
        if self.sysClr is not None:
            self.remove(self.sysClr)
        if self.srgbClr is not None:
            self.remove(self.srgbClr)
        # Add new srgbClr element - uses xmlchemy magic _add_ method
        srgb_clr: CT_SRgbColor = self._add_srgbClr()  # pyright: ignore[reportAttributeAccessIssue,reportUnknownMemberType,reportUnknownVariableType]
        srgb_clr.val = value.upper()


class CT_SystemColor(BaseOxmlElement):
    """``<a:sysClr>`` element, represents a system-defined color."""

    val: str | None = OptionalAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]
    lastClr: str | None = OptionalAttribute("lastClr", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_SRgbColor(BaseOxmlElement):
    """``<a:srgbClr>`` element, represents an RGB color."""

    val: str | None = OptionalAttribute("val", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_FontScheme(BaseOxmlElement):
    """``<a:fontScheme>`` element.

    Defines the major and minor fonts for the theme.
    """

    name: str | None = OptionalAttribute("name", XsdString)  # pyright: ignore[reportAssignmentType]

    majorFont: CT_FontCollection | None = ZeroOrOne("a:majorFont")  # pyright: ignore[reportAssignmentType]
    minorFont: CT_FontCollection | None = ZeroOrOne("a:minorFont")  # pyright: ignore[reportAssignmentType]


class CT_FontCollection(BaseOxmlElement):
    """``<a:majorFont>`` or ``<a:minorFont>`` element.

    Contains font definitions for Latin, East Asian, and Complex Script text.
    """

    latin: CT_TextFont | None = ZeroOrOne("a:latin")  # pyright: ignore[reportAssignmentType]
    ea: CT_TextFont | None = ZeroOrOne("a:ea")  # pyright: ignore[reportAssignmentType]
    cs: CT_TextFont | None = ZeroOrOne("a:cs")  # pyright: ignore[reportAssignmentType]


class CT_TextFont(BaseOxmlElement):
    """``<a:latin>``, ``<a:ea>``, or ``<a:cs>`` element.

    Specifies a font typeface for a particular script.
    """

    typeface: str | None = OptionalAttribute("typeface", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_FontCollectionExtensionList(BaseOxmlElement):
    """``<a:extLst>`` element within font collection - placeholder for future extensions."""

    pass
