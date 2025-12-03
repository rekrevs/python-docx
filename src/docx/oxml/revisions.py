"""Custom element classes for revisions (track changes)."""

from __future__ import annotations

from datetime import datetime
from typing import TYPE_CHECKING, List

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, RequiredAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.text.run import CT_R


class CT_TrackChange(BaseOxmlElement):
    """Base class for track change elements (`w:ins`, `w:del`, etc.).

    Track changes record insertions, deletions, and formatting changes
    made to a document when revision tracking is enabled.

    Common attributes:
        w:id - Unique identifier for this revision
        w:author - Name of the person who made the change
        w:date - Date/time when the change was made (ISO 8601 format)
    """

    id: int = RequiredAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    author: str = OptionalAttribute("w:author", ST_String)  # pyright: ignore[reportAssignmentType]
    date: str | None = OptionalAttribute("w:date", ST_String)  # pyright: ignore[reportAssignmentType]

    @property
    def revision_id(self) -> int:
        """The unique identifier for this revision."""
        return self.id

    @property
    def revision_author(self) -> str:
        """The author who made this change."""
        return self.author or ""

    @property
    def revision_date(self) -> datetime | None:
        """The date/time when this change was made, or None if not available.

        The date is parsed from ISO 8601 format (e.g., "2025-09-23T12:09:00Z").
        """
        if self.date is None:
            return None
        try:
            # Handle both "Z" suffix and without
            date_str = self.date.rstrip("Z")
            return datetime.fromisoformat(date_str)
        except (ValueError, TypeError):
            return None


class CT_RunTrackChange(CT_TrackChange):
    """`w:ins` or `w:del` element containing tracked run-level changes.

    These elements wrap runs (`w:r`) that have been inserted or deleted.
    The content inside represents what was added or removed.

    Example XML (insertion):
        <w:ins w:id="1" w:author="John" w:date="2025-01-15T10:30:00Z">
            <w:r><w:t>inserted text</w:t></w:r>
        </w:ins>

    Example XML (deletion):
        <w:del w:id="2" w:author="Jane" w:date="2025-01-15T11:00:00Z">
            <w:r><w:delText>deleted text</w:delText></w:r>
        </w:del>
    """

    # -- children --
    r = ZeroOrMore("w:r", successors=())

    # -- type-declarations for methods added by metaclass --
    r_lst: List[CT_R]

    @property
    def revision_text(self) -> str:
        """The text content of this tracked change.

        For insertions, returns the inserted text.
        For deletions, returns the deleted text (from w:delText elements).
        """
        # Check if this is a deletion (look for delText) or insertion (look for t)
        del_texts = self.xpath(".//w:delText/text()")
        if del_texts:
            return "".join(del_texts)
        # Otherwise get regular text
        texts = self.xpath(".//w:t/text()")
        return "".join(texts)

    @property
    def is_insertion(self) -> bool:
        """True if this is an insertion (w:ins) element."""
        from docx.oxml.ns import qn
        return self.tag == qn("w:ins")

    @property
    def is_deletion(self) -> bool:
        """True if this is a deletion (w:del) element."""
        from docx.oxml.ns import qn
        return self.tag == qn("w:del")

    def accept(self) -> None:
        """Accept this revision, applying the change and removing markup.

        For insertions: moves content outside the w:ins and removes the w:ins element.
        For deletions: removes the w:del element (keeping the text deleted).
        """
        parent = self.getparent()
        if parent is None:
            return

        if self.is_insertion:
            # Move all children (runs) to before the ins element
            for child in list(self):
                self.addprevious(child)
            # Remove the ins element
            parent.remove(self)
        elif self.is_deletion:
            # Just remove the del element (text stays deleted)
            parent.remove(self)

    def reject(self) -> None:
        """Reject this revision, reverting the change and removing markup.

        For insertions: removes the w:ins element and all its content.
        For deletions: converts w:delText to w:t, moves runs outside w:del, removes w:del.
        """
        parent = self.getparent()
        if parent is None:
            return

        if self.is_insertion:
            # Remove the entire ins element with its content
            parent.remove(self)
        elif self.is_deletion:
            # Convert delText elements to regular text elements
            for del_text in self.xpath(".//w:delText"):
                # Create a new w:t element with the same text
                from docx.oxml.parser import OxmlElement
                t_elem = OxmlElement("w:t")
                t_elem.text = del_text.text
                # Preserve xml:space attribute if present
                space_attr = del_text.get("{http://www.w3.org/XML/1998/namespace}space")
                if space_attr:
                    t_elem.set("{http://www.w3.org/XML/1998/namespace}space", space_attr)
                # Replace delText with t
                del_text_parent = del_text.getparent()
                if del_text_parent is not None:
                    del_text_parent.replace(del_text, t_elem)

            # Move all children (runs) to before the del element
            for child in list(self):
                self.addprevious(child)
            # Remove the del element
            parent.remove(self)


class CT_PPrChange(CT_TrackChange):
    """`w:pPrChange` element recording paragraph property changes.

    This element appears inside `w:pPr` and records the previous
    paragraph formatting before a change was made.
    """
    pass


class CT_RPrChange(CT_TrackChange):
    """`w:rPrChange` element recording run property (formatting) changes.

    This element appears inside `w:rPr` and records the previous
    character formatting before a change was made.
    """
    pass


class CT_SectPrChange(CT_TrackChange):
    """`w:sectPrChange` element recording section property changes."""
    pass


class CT_TblPrChange(CT_TrackChange):
    """`w:tblPrChange` element recording table property changes."""
    pass


class CT_TcPrChange(CT_TrackChange):
    """`w:tcPrChange` element recording table cell property changes."""
    pass


class CT_TrPrChange(CT_TrackChange):
    """`w:trPrChange` element recording table row property changes."""
    pass
