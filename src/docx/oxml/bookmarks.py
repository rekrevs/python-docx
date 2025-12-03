"""Custom element classes for bookmarks."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute

if TYPE_CHECKING:
    pass


class CT_Bookmark(BaseOxmlElement):
    """`w:bookmarkStart` element, marking the start of a bookmark.

    A bookmark is a named range in a document that can be used as a target
    for cross-references, hyperlinks, or programmatic access to specific
    document locations.

    Bookmarks are delimited by `w:bookmarkStart` and `w:bookmarkEnd` elements
    with matching `w:id` attributes.

    Example XML:
        <w:bookmarkStart w:id="0" w:name="_Toc123456789"/>
        ... content within bookmark ...
        <w:bookmarkEnd w:id="0"/>
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    name: str = RequiredAttribute("w:name", ST_String)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, bookmark_id: int, bookmark_name: str) -> CT_Bookmark:
        """Create a new `w:bookmarkStart` element.

        Args:
            bookmark_id: The unique numeric id for this bookmark.
            bookmark_name: The name for this bookmark.

        Returns:
            A new CT_Bookmark (bookmarkStart) element.
        """
        from docx.oxml.parser import OxmlElement

        bookmark_start = OxmlElement("w:bookmarkStart")
        bookmark_start.set(qn("w:id"), str(bookmark_id))
        bookmark_start.set(qn("w:name"), bookmark_name)
        return cast("CT_Bookmark", bookmark_start)

    @property
    def bookmark_id(self) -> int:
        """The numeric id of this bookmark.

        This id links the bookmarkStart to its corresponding bookmarkEnd.
        """
        return self.id

    @property
    def bookmark_name(self) -> str:
        """The name of this bookmark.

        Bookmark names are used for cross-references and programmatic access.
        Common naming conventions:
        - `_Toc*`: Table of contents entries
        - `_Ref*`: Cross-reference targets
        - `_Hlk*`: Hyperlink targets
        - Custom names for user-defined bookmarks
        """
        return self.name


class CT_MarkupRange(BaseOxmlElement):
    """`w:bookmarkEnd` element, marking the end of a bookmark.

    This element contains only the `w:id` attribute which links it to
    the corresponding `w:bookmarkStart` element.
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, bookmark_id: int) -> CT_MarkupRange:
        """Create a new `w:bookmarkEnd` element.

        Args:
            bookmark_id: The numeric id matching the corresponding bookmarkStart.

        Returns:
            A new CT_MarkupRange (bookmarkEnd) element.
        """
        from docx.oxml.parser import OxmlElement

        bookmark_end = OxmlElement("w:bookmarkEnd")
        bookmark_end.set(qn("w:id"), str(bookmark_id))
        return cast("CT_MarkupRange", bookmark_end)

    @property
    def bookmark_id(self) -> int:
        """The numeric id linking this end marker to its bookmarkStart."""
        return self.id
