"""Custom element classes for footnotes and endnotes."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P


class CT_Footnote(BaseOxmlElement):
    """`w:footnote` element, containing a single footnote.

    A footnote contains one or more paragraphs of content and has an id attribute
    that links it to a footnote reference in the document body.
    """

    # -- children --
    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --
    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    # -- attributes --
    id: int = OptionalAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", ST_String)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, footnote_id: int, text: str = "", footnote_type: str | None = None) -> CT_Footnote:
        """Create a new `w:footnote` element.

        Args:
            footnote_id: The unique id for this footnote.
            text: Optional text content for the first paragraph.
            footnote_type: Optional type ('separator', 'continuationSeparator', or None for normal).

        Returns:
            A new CT_Footnote element with a paragraph containing the footnote marker and text.
        """
        from docx.oxml.parser import OxmlElement

        footnote = OxmlElement("w:footnote")
        footnote.set(qn("w:id"), str(footnote_id))

        if footnote_type:
            footnote.set(qn("w:type"), footnote_type)

        # Create paragraph with footnote style
        p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), "FootnoteText")
        pPr.append(pStyle)
        p.append(pPr)

        # For normal footnotes, add footnote reference marker and text
        if footnote_type is None:
            # Add run with footnote reference (the auto-numbering marker)
            r_ref = OxmlElement("w:r")
            rPr_ref = OxmlElement("w:rPr")
            rStyle_ref = OxmlElement("w:rStyle")
            rStyle_ref.set(qn("w:val"), "FootnoteReference")
            rPr_ref.append(rStyle_ref)
            r_ref.append(rPr_ref)
            footnoteRef = OxmlElement("w:footnoteRef")
            r_ref.append(footnoteRef)
            p.append(r_ref)

            # Add run with text content
            if text:
                r_text = OxmlElement("w:r")
                t = OxmlElement("w:t")
                t.text = " " + text  # Space between marker and text
                r_text.append(t)
                p.append(r_text)
        else:
            # For separator footnotes, just add a separator run
            r = OxmlElement("w:r")
            separator = OxmlElement("w:separator") if footnote_type == "separator" else OxmlElement("w:continuationSeparator")
            r.append(separator)
            p.append(r)

        footnote.append(p)
        return cast("CT_Footnote", footnote)

    @property
    def footnote_id(self) -> int:
        """The id of this footnote."""
        return self.id

    @property
    def footnote_type(self) -> str | None:
        """The type of this footnote (separator, continuationSeparator, or None for normal)."""
        return self.type

    @property
    def is_separator(self) -> bool:
        """True if this is a separator footnote (type='separator' or 'continuationSeparator')."""
        return self.type in ("separator", "continuationSeparator")


class CT_Endnote(BaseOxmlElement):
    """`w:endnote` element, containing a single endnote.

    An endnote contains one or more paragraphs of content and has an id attribute
    that links it to an endnote reference in the document body.
    """

    # -- children --
    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --
    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    # -- attributes --
    id: int = OptionalAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]
    type: str | None = OptionalAttribute("w:type", ST_String)  # pyright: ignore[reportAssignmentType]

    @property
    def endnote_id(self) -> int:
        """The id of this endnote."""
        return self.id

    @property
    def endnote_type(self) -> str | None:
        """The type of this endnote (separator, continuationSeparator, or None for normal)."""
        return self.type

    @property
    def is_separator(self) -> bool:
        """True if this is a separator endnote (type='separator' or 'continuationSeparator')."""
        return self.type in ("separator", "continuationSeparator")


class CT_Footnotes(BaseOxmlElement):
    """`w:footnotes` element, the root element of the footnotes part.

    Contains all footnotes in the document, including separator footnotes
    (used for rendering the line between body text and footnotes).
    """

    footnote = ZeroOrMore("w:footnote", successors=())

    @classmethod
    def new(cls) -> CT_Footnotes:
        """Create a new `w:footnotes` element with separator footnotes.

        Returns:
            A new CT_Footnotes element containing the required separator footnotes
            (id=-1 for separator, id=0 for continuation separator).
        """
        from docx.oxml.parser import OxmlElement

        footnotes = OxmlElement("w:footnotes")

        # Add separator footnote (id=-1)
        separator = CT_Footnote.new(-1, footnote_type="separator")
        footnotes.append(separator)

        # Add continuation separator footnote (id=0)
        cont_separator = CT_Footnote.new(0, footnote_type="continuationSeparator")
        footnotes.append(cont_separator)

        return cast("CT_Footnotes", footnotes)

    @property
    def footnote_lst(self) -> List[CT_Footnote]:
        """List of all footnote elements."""
        return self.xpath("./w:footnote")

    def footnote_by_id(self, footnote_id: int) -> CT_Footnote | None:
        """Return the footnote with the specified id, or None if not found."""
        for fn in self.footnote_lst:
            if fn.footnote_id == footnote_id:
                return fn
        return None

    def add_footnote(self, text: str = "") -> CT_Footnote:
        """Add a new footnote to this container.

        Args:
            text: The text content for the footnote.

        Returns:
            The newly created CT_Footnote element.
        """
        # Find next available ID (max ID + 1, starting from 1)
        max_id = max((fn.footnote_id for fn in self.footnote_lst if fn.footnote_id >= 0), default=0)
        next_id = max_id + 1

        footnote = CT_Footnote.new(next_id, text)
        self.append(footnote)
        return footnote


class CT_Endnotes(BaseOxmlElement):
    """`w:endnotes` element, the root element of the endnotes part.

    Contains all endnotes in the document, including separator endnotes.
    """

    endnote = ZeroOrMore("w:endnote", successors=())

    @property
    def endnote_lst(self) -> List[CT_Endnote]:
        """List of all endnote elements."""
        return self.xpath("./w:endnote")

    def endnote_by_id(self, endnote_id: int) -> CT_Endnote | None:
        """Return the endnote with the specified id, or None if not found."""
        for en in self.endnote_lst:
            if en.endnote_id == endnote_id:
                return en
        return None


class CT_FootnoteReference(BaseOxmlElement):
    """`w:footnoteReference` element, a reference to a footnote in the document body."""

    id: int = OptionalAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, footnote_id: int) -> CT_FootnoteReference:
        """Create a new `w:footnoteReference` element.

        Args:
            footnote_id: The id of the footnote to reference.

        Returns:
            A new CT_FootnoteReference element.
        """
        from docx.oxml.parser import OxmlElement

        ref = OxmlElement("w:footnoteReference")
        ref.set(qn("w:id"), str(footnote_id))
        return cast("CT_FootnoteReference", ref)

    @property
    def footnote_id(self) -> int:
        """The id of the referenced footnote."""
        return self.id


class CT_EndnoteReference(BaseOxmlElement):
    """`w:endnoteReference` element, a reference to an endnote in the document body."""

    id: int = OptionalAttribute("w:id", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    @property
    def endnote_id(self) -> int:
        """The id of the referenced endnote."""
        return self.id
