"""Custom element classes related to Structured Document Tags (Content Controls)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List

from docx.oxml.simpletypes import ST_DecimalNumber, ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.oxml.text.run import CT_R


class CT_SdtPr(BaseOxmlElement):
    """`w:sdtPr` element, containing properties for a content control.

    This element specifies the set of properties for a single structured document tag.
    """

    # -- Child elements that identify the SDT type --
    # Only one of these should be present to indicate the type
    text: CT_SdtText | None = ZeroOrOne("w:text", successors=())  # pyright: ignore[reportAssignmentType]
    date: CT_SdtDate | None = ZeroOrOne("w:date", successors=())  # pyright: ignore[reportAssignmentType]
    dropDownList: CT_SdtDropDownList | None = ZeroOrOne("w:dropDownList", successors=())  # pyright: ignore[reportAssignmentType]
    comboBox: CT_SdtComboBox | None = ZeroOrOne("w:comboBox", successors=())  # pyright: ignore[reportAssignmentType]
    docPartObj: CT_SdtDocPartObj | None = ZeroOrOne("w:docPartObj", successors=())  # pyright: ignore[reportAssignmentType]
    picture: BaseOxmlElement | None = ZeroOrOne("w:picture", successors=())  # pyright: ignore[reportAssignmentType]

    @property
    def tag_val(self) -> str | None:
        """The value of the w:tag child element's w:val attribute, or None if not present."""
        tag_elements = self.xpath("./w:tag/@w:val")
        return tag_elements[0] if tag_elements else None

    @tag_val.setter
    def tag_val(self, value: str | None) -> None:
        """Set or remove the w:tag element."""
        self._remove_tag()
        if value is not None:
            from docx.oxml.ns import nsdecls
            from docx.oxml.parser import parse_xml

            tag_elm = parse_xml(f'<w:tag {nsdecls("w")} w:val="{value}"/>')
            self.insert(0, tag_elm)

    def _remove_tag(self) -> None:
        """Remove the w:tag child element if present."""
        tag_elements = self.xpath("./w:tag")
        for tag in tag_elements:
            self.remove(tag)

    @property
    def alias_val(self) -> str | None:
        """The value of the w:alias child element's w:val attribute, or None if not present."""
        alias_elements = self.xpath("./w:alias/@w:val")
        return alias_elements[0] if alias_elements else None

    @alias_val.setter
    def alias_val(self, value: str | None) -> None:
        """Set or remove the w:alias element."""
        self._remove_alias()
        if value is not None:
            from docx.oxml.ns import nsdecls
            from docx.oxml.parser import parse_xml

            alias_elm = parse_xml(f'<w:alias {nsdecls("w")} w:val="{value}"/>')
            self.insert(0, alias_elm)

    def _remove_alias(self) -> None:
        """Remove the w:alias child element if present."""
        alias_elements = self.xpath("./w:alias")
        for alias in alias_elements:
            self.remove(alias)

    @property
    def sdt_id(self) -> int | None:
        """The value of the w:id child element's w:val attribute, or None if not present."""
        id_elements = self.xpath("./w:id/@w:val")
        return int(id_elements[0]) if id_elements else None

    @property
    def sdt_type(self) -> str:
        """Return a string identifying the type of this content control.

        Returns one of: 'text', 'date', 'dropDownList', 'comboBox', 'docPartObj',
        'picture', 'richText', or 'unknown'.
        """
        if self.text is not None:
            return "text"
        if self.date is not None:
            return "date"
        if self.dropDownList is not None:
            return "dropDownList"
        if self.comboBox is not None:
            return "comboBox"
        if self.docPartObj is not None:
            return "docPartObj"
        if self.picture is not None:
            return "picture"
        # Rich text SDT has no type-specific child element
        # Check if there's content that looks like rich text
        return "richText"


class CT_SdtText(BaseOxmlElement):
    """`w:text` element, indicating a plain text content control."""

    multiLine: bool | None = OptionalAttribute("w:multiLine", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_SdtDate(BaseOxmlElement):
    """`w:date` element, indicating a date picker content control."""

    fullDate: str | None = OptionalAttribute("w:fullDate", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_SdtListItem(BaseOxmlElement):
    """`w:listItem` element, an item in a dropdown or combo box."""

    displayText: str | None = OptionalAttribute("w:displayText", ST_String)  # pyright: ignore[reportAssignmentType]
    value: str | None = OptionalAttribute("w:value", ST_String)  # pyright: ignore[reportAssignmentType]


class CT_SdtDropDownList(BaseOxmlElement):
    """`w:dropDownList` element, indicating a dropdown list content control."""

    listItem = ZeroOrMore("w:listItem", successors=())

    listItem_lst: List[CT_SdtListItem]


class CT_SdtComboBox(BaseOxmlElement):
    """`w:comboBox` element, indicating a combo box content control."""

    listItem = ZeroOrMore("w:listItem", successors=())

    listItem_lst: List[CT_SdtListItem]


class CT_SdtDocPartObj(BaseOxmlElement):
    """`w:docPartObj` element, indicating a document part object (e.g., TOC)."""

    pass


class CT_SdtContentBlock(BaseOxmlElement):
    """`w:sdtContent` element for block-level content controls.

    This element contains the content of a block-level structured document tag,
    which can include paragraphs and tables.
    """

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    # -- type-declarations for methods added by metaclass --
    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` elements in this content, in document order."""
        return self.xpath("./w:p | ./w:tbl")


class CT_SdtContentRun(BaseOxmlElement):
    """`w:sdtContent` element for inline/run-level content controls.

    This element contains the content of an inline structured document tag,
    which contains runs.
    """

    r = ZeroOrMore("w:r", successors=())

    r_lst: List[CT_R]

    @property
    def text(self) -> str:
        """The text content of all runs in this inline content control."""
        return "".join(r.text for r in self.r_lst)


class CT_SdtBlock(BaseOxmlElement):
    """`w:sdt` element for block-level content controls.

    A block-level structured document tag can appear as a sibling to paragraphs
    and tables in the document body, table cells, headers, footers, etc.
    """

    sdtPr: CT_SdtPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtPr", successors=("w:sdtEndPr", "w:sdtContent")
    )
    sdtContent: CT_SdtContentBlock | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtContent", successors=()
    )

    # -- type-declarations for methods added by metaclass --
    get_or_add_sdtPr: Callable[[], CT_SdtPr]
    get_or_add_sdtContent: Callable[[], CT_SdtContentBlock]

    # -- delegated content properties --
    @property
    def p_lst(self) -> List[CT_P]:
        """All `w:p` elements in this content control's content."""
        content = self.sdtContent
        return content.p_lst if content is not None else []

    @property
    def tbl_lst(self) -> List[CT_Tbl]:
        """All `w:tbl` elements in this content control's content."""
        content = self.sdtContent
        return content.tbl_lst if content is not None else []

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` elements in this content control, in document order."""
        content = self.sdtContent
        return content.inner_content_elements if content is not None else []

    def add_p(self) -> CT_P:
        """Add a new `w:p` element to this content control's content."""
        content = self.get_or_add_sdtContent()
        return content.add_p()

    def _insert_tbl(self, tbl: CT_Tbl) -> CT_Tbl:
        """Insert a `w:tbl` element into this content control's content."""
        content = self.get_or_add_sdtContent()
        return content._insert_tbl(tbl)

    @property
    def sdt_tag(self) -> str | None:
        """The tag value of this content control, or None if not set."""
        pr = self.sdtPr
        return pr.tag_val if pr is not None else None

    @sdt_tag.setter
    def sdt_tag(self, value: str | None) -> None:
        """Set the tag value of this content control."""
        pr = self.get_or_add_sdtPr()
        pr.tag_val = value

    @property
    def alias(self) -> str | None:
        """The alias (title) of this content control, or None if not set."""
        pr = self.sdtPr
        return pr.alias_val if pr is not None else None

    @alias.setter
    def alias(self, value: str | None) -> None:
        """Set the alias (title) of this content control."""
        pr = self.get_or_add_sdtPr()
        pr.alias_val = value

    @property
    def sdt_type(self) -> str:
        """Return the type of this content control."""
        pr = self.sdtPr
        return pr.sdt_type if pr is not None else "richText"


class CT_SdtRun(BaseOxmlElement):
    """`w:sdt` element for inline/run-level content controls.

    An inline structured document tag appears within a paragraph, containing runs
    rather than block-level content.
    """

    sdtPr: CT_SdtPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtPr", successors=("w:sdtEndPr", "w:sdtContent")
    )
    sdtContent: CT_SdtContentRun | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:sdtContent", successors=()
    )

    # -- type-declarations for methods added by metaclass --
    get_or_add_sdtPr: Callable[[], CT_SdtPr]
    get_or_add_sdtContent: Callable[[], CT_SdtContentRun]

    @property
    def text(self) -> str:
        """The text content of this inline content control."""
        content = self.sdtContent
        return content.text if content is not None else ""

    @property
    def sdt_tag(self) -> str | None:
        """The tag value of this content control, or None if not set."""
        pr = self.sdtPr
        return pr.tag_val if pr is not None else None

    @property
    def alias(self) -> str | None:
        """The alias (title) of this content control, or None if not set."""
        pr = self.sdtPr
        return pr.alias_val if pr is not None else None

    @property
    def sdt_type(self) -> str:
        """Return the type of this content control."""
        pr = self.sdtPr
        return pr.sdt_type if pr is not None else "richText"
