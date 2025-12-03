"""|FootnotesPart| and |EndnotesPart| and closely related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, List

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.opc.package import OpcPackage
    from docx.oxml.footnotes import CT_Endnote, CT_Endnotes, CT_Footnote, CT_Footnotes


class FootnotesPart(XmlPart):
    """Proxy for the footnotes.xml part containing footnotes for a document."""

    @classmethod
    def default(cls, package: OpcPackage) -> FootnotesPart:
        """Return a newly created footnotes part with separator footnotes.

        Args:
            package: The package this part belongs to.

        Returns:
            A new FootnotesPart with the required separator footnotes.
        """
        from docx.oxml.footnotes import CT_Footnotes

        partname = PackURI("/word/footnotes.xml")
        content_type = CT.WML_FOOTNOTES
        element = CT_Footnotes.new()
        return cls(partname, content_type, element, package)

    @lazyproperty
    def footnotes(self) -> List[CT_Footnote]:
        """List of all footnote elements in this part.

        Includes separator footnotes which are used for rendering but typically
        filtered out when iterating user-visible footnotes.
        """
        return self._element.footnote_lst

    def footnote_by_id(self, footnote_id: int) -> CT_Footnote | None:
        """Return the footnote with the specified id, or None if not found."""
        return self._element.footnote_by_id(footnote_id)

    def add_footnote(self, text: str = "") -> CT_Footnote:
        """Add a new footnote to this part.

        Args:
            text: The text content for the footnote.

        Returns:
            The newly created CT_Footnote element.
        """
        # Clear the cached footnotes list since we're adding a new one
        if "footnotes" in self.__dict__:
            del self.__dict__["footnotes"]
        return self._element.add_footnote(text)


class EndnotesPart(XmlPart):
    """Proxy for the endnotes.xml part containing endnotes for a document."""

    @lazyproperty
    def endnotes(self) -> List[CT_Endnote]:
        """List of all endnote elements in this part.

        Includes separator endnotes which are used for rendering but typically
        filtered out when iterating user-visible endnotes.
        """
        return self._element.endnote_lst

    def endnote_by_id(self, endnote_id: int) -> CT_Endnote | None:
        """Return the endnote with the specified id, or None if not found."""
        return self._element.endnote_by_id(endnote_id)
