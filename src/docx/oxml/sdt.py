"""Custom element classes related to Structured Document Tags (Content Controls)."""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_String
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
    # Note: Named "text_elm" to avoid conflict with lxml's text property
    text_elm: CT_SdtText | None = ZeroOrOne("w:text", successors=())  # pyright: ignore[reportAssignmentType]
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
        if self.text_elm is not None:
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

    is_multiLine: bool | None = OptionalAttribute("w:multiLine", ST_String)  # pyright: ignore[reportAssignmentType]


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

    def add_list_item(self, display_text: str, value: str | None = None) -> CT_SdtListItem:
        """Add a list item to this dropdown.

        Args:
            display_text: The text displayed to the user.
            value: The value stored when selected. Defaults to display_text if not provided.

        Returns:
            The newly created CT_SdtListItem element.
        """
        from docx.oxml.parser import OxmlElement

        item = OxmlElement("w:listItem")
        item.set(qn("w:displayText"), display_text)
        item.set(qn("w:value"), value if value is not None else display_text)
        self.append(item)
        return cast(CT_SdtListItem, item)


class CT_SdtComboBox(BaseOxmlElement):
    """`w:comboBox` element, indicating a combo box content control."""

    listItem = ZeroOrMore("w:listItem", successors=())

    listItem_lst: List[CT_SdtListItem]

    def add_list_item(self, display_text: str, value: str | None = None) -> CT_SdtListItem:
        """Add a list item to this combo box.

        Args:
            display_text: The text displayed to the user.
            value: The value stored when selected. Defaults to display_text if not provided.

        Returns:
            The newly created CT_SdtListItem element.
        """
        from docx.oxml.parser import OxmlElement

        item = OxmlElement("w:listItem")
        item.set(qn("w:displayText"), display_text)
        item.set(qn("w:value"), value if value is not None else display_text)
        self.append(item)
        return cast(CT_SdtListItem, item)


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
    def sdt_text(self) -> str:
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

    @classmethod
    def new(
        cls,
        sdt_type: str = "richText",
        tag: str | None = None,
        alias: str | None = None,
        placeholder_text: str = "",
    ) -> CT_SdtBlock:
        """Create a new block-level content control.

        Args:
            sdt_type: The type of content control. One of 'richText', 'text', 'date',
                     'dropDownList', 'comboBox'. Default is 'richText'.
            tag: Optional tag value for identifying the control programmatically.
            alias: Optional alias (title) displayed in the UI.
            placeholder_text: Optional placeholder text for the content.

        Returns:
            A new CT_SdtBlock element.
        """
        from docx.oxml.parser import OxmlElement

        sdt = OxmlElement("w:sdt")

        # Create sdtPr
        sdtPr = OxmlElement("w:sdtPr")

        # Add alias if provided
        if alias:
            alias_elm = OxmlElement("w:alias")
            alias_elm.set(qn("w:val"), alias)
            sdtPr.append(alias_elm)

        # Add tag if provided
        if tag:
            tag_elm = OxmlElement("w:tag")
            tag_elm.set(qn("w:val"), tag)
            sdtPr.append(tag_elm)

        # Add type-specific element
        if sdt_type == "text":
            text_elm = OxmlElement("w:text")
            sdtPr.append(text_elm)
        elif sdt_type == "date":
            date_elm = OxmlElement("w:date")
            sdtPr.append(date_elm)
        elif sdt_type == "dropDownList":
            ddl_elm = OxmlElement("w:dropDownList")
            sdtPr.append(ddl_elm)
        elif sdt_type == "comboBox":
            combo_elm = OxmlElement("w:comboBox")
            sdtPr.append(combo_elm)
        # richText has no type-specific element

        sdt.append(sdtPr)

        # Create sdtContent with a paragraph
        sdtContent = OxmlElement("w:sdtContent")
        p = OxmlElement("w:p")
        if placeholder_text:
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = placeholder_text
            r.append(t)
            p.append(r)
        sdtContent.append(p)
        sdt.append(sdtContent)

        return cast("CT_SdtBlock", sdt)

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
        return content._insert_tbl(tbl)  # pyright: ignore[reportPrivateUsage]

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

    def add_list_item(self, display_text: str, value: str | None = None) -> CT_SdtListItem:
        """Add a list item to this content control (for dropDownList or comboBox types).

        Args:
            display_text: The text displayed to the user.
            value: The value stored when selected. Defaults to display_text if not provided.

        Returns:
            The newly created CT_SdtListItem element.

        Raises:
            ValueError: If this content control is not a dropDownList or comboBox.
        """
        pr = self.get_or_add_sdtPr()

        if pr.dropDownList is not None:
            return pr.dropDownList.add_list_item(display_text, value)
        elif pr.comboBox is not None:
            return pr.comboBox.add_list_item(display_text, value)
        else:
            raise ValueError(
                f"Cannot add list item to content control of type '{self.sdt_type}'. "
                "Only dropDownList and comboBox types support list items."
            )


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
    def sdt_text(self) -> str:
        """The text content of this inline content control."""
        content = self.sdtContent
        return content.sdt_text if content is not None else ""

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
