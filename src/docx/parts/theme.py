"""Provides ThemePart and related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.opc.part import XmlPart

if TYPE_CHECKING:
    from docx.oxml.theme import CT_OfficeStyleSheet


class ThemePart(XmlPart):
    """Proxy for the theme.xml part containing theme definitions for a document.

    The theme part defines the color scheme, font scheme, and format scheme (effects)
    that are used to style the document.
    """

    @property
    def theme_element(self) -> CT_OfficeStyleSheet:
        """Return the root ``<a:theme>`` element for this theme part."""
        return self.element
