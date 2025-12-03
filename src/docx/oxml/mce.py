"""Custom element classes for Markup Compatibility and Extensibility (MCE).

MCE is defined in ECMA-376 Part 3 / ISO 29500-3 and provides mechanisms for
forward compatibility in OOXML documents.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List

from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_AlternateContent(BaseOxmlElement):
    """`mc:AlternateContent` element.

    Container for version-specific markup alternatives. Contains one or more
    Choice elements followed by an optional Fallback element.
    """

    @property
    def choice(self) -> CT_Choice | None:
        """First `mc:Choice` child element, or None if not present."""
        choices = self.xpath("mc:Choice")
        return choices[0] if choices else None

    @property
    def choices(self) -> List[CT_Choice]:
        """All `mc:Choice` child elements."""
        return self.xpath("mc:Choice")

    @property
    def fallback(self) -> CT_Fallback | None:
        """`mc:Fallback` child element, or None if not present."""
        fallbacks = self.xpath("mc:Fallback")
        return fallbacks[0] if fallbacks else None

    @property
    def txbxContent_choice(self) -> CT_TxbxContent | None:
        """The w:txbxContent element from mc:Choice, if this is a text box."""
        matches = self.xpath("mc:Choice//wps:txbx/w:txbxContent")
        return matches[0] if matches else None

    @property
    def txbxContent_fallback(self) -> CT_TxbxContent | None:
        """The w:txbxContent element from mc:Fallback, if this is a text box."""
        matches = self.xpath("mc:Fallback//v:textbox/w:txbxContent")
        return matches[0] if matches else None

    @property
    def is_textbox(self) -> bool:
        """True if this AlternateContent contains a text box."""
        return self.txbxContent_choice is not None or self.txbxContent_fallback is not None


class CT_Choice(BaseOxmlElement):
    """`mc:Choice` element.

    Contains content for applications that understand the namespaces
    specified in the Requires attribute.
    """

    @property
    def requires(self) -> str:
        """The Requires attribute value (space-separated namespace prefixes)."""
        return self.get("Requires", "")

    @property
    def requires_namespaces(self) -> List[str]:
        """List of namespace prefixes required by this Choice."""
        req = self.requires
        return req.split() if req else []


class CT_Fallback(BaseOxmlElement):
    """`mc:Fallback` element.

    Contains fallback content for applications that don't understand
    the namespaces required by any Choice element.
    """
    pass


class CT_TxbxContent(BaseOxmlElement):
    """`w:txbxContent` element.

    Container for text box content. Can contain paragraphs and tables.
    This element appears in both mc:Choice (inside wps:txbx) and
    mc:Fallback (inside v:textbox).
    """

    add_p: Callable[[], CT_P]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    @property
    def p_lst(self) -> List[CT_P]:
        """All `w:p` (paragraph) elements in this text box content."""
        return self.xpath("w:p")

    @property
    def tbl_lst(self) -> List[CT_Tbl]:
        """All `w:tbl` (table) elements in this text box content."""
        return self.xpath("w:tbl")

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` elements in document order."""
        return self.xpath("./w:p | ./w:tbl")
