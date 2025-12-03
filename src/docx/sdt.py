"""Structured Document Tags (Content Controls) proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    import docx.types as t
    from docx.oxml.sdt import CT_SdtBlock, CT_SdtRun
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph


class SdtContentControls:
    """Collection of content controls (SDT elements) in a container.

    This collection provides access to all structured document tags in a document body,
    table cell, header, footer, or other block container.
    """

    def __init__(
        self, sdt_elements: list[CT_SdtBlock], parent: t.ProvidesStoryPart
    ) -> None:
        self._sdt_elements = sdt_elements
        self._parent = parent

    def __iter__(self) -> Iterator[SdtBlockContentControl]:
        """Iterate over content controls in this collection."""
        return (
            SdtBlockContentControl(sdt_elm, self._parent)
            for sdt_elm in self._sdt_elements
        )

    def __len__(self) -> int:
        """Return the number of content controls in this collection."""
        return len(self._sdt_elements)

    def __getitem__(self, index: int) -> SdtBlockContentControl:
        """Return the content control at `index`."""
        return SdtBlockContentControl(self._sdt_elements[index], self._parent)

    def get_by_tag(self, tag: str) -> SdtBlockContentControl | None:
        """Return the first content control with the specified tag, or None.

        The tag is a user-defined string that can be used to identify content controls.
        """
        for sdt_elm in self._sdt_elements:
            if sdt_elm.sdt_tag == tag:
                return SdtBlockContentControl(sdt_elm, self._parent)
        return None

    def get_by_alias(self, alias: str) -> SdtBlockContentControl | None:
        """Return the first content control with the specified alias (title), or None.

        The alias is displayed in the Word UI as the content control's title.
        """
        for sdt_elm in self._sdt_elements:
            if sdt_elm.alias == alias:
                return SdtBlockContentControl(sdt_elm, self._parent)
        return None

    def filter_by_type(self, sdt_type: str) -> Iterator[SdtBlockContentControl]:
        """Yield content controls of the specified type.

        `sdt_type` should be one of: 'text', 'date', 'dropDownList', 'comboBox',
        'docPartObj', 'picture', 'richText'.
        """
        for sdt_elm in self._sdt_elements:
            if sdt_elm.sdt_type == sdt_type:
                yield SdtBlockContentControl(sdt_elm, self._parent)


class SdtBlockContentControl(BlockItemContainer):
    """Proxy for a block-level structured document tag (content control).

    A block-level content control can contain paragraphs and tables, similar to
    a table cell. It appears at the same level as paragraphs in the document body,
    table cells, headers, footers, etc.

    Content controls are used for:
    - Fillable form fields (text boxes, dropdowns, date pickers, checkboxes)
    - Repeating content regions
    - Document building blocks (like table of contents)
    - Data binding to XML data stores

    Common types include:
    - Plain text: Single-line or multi-line text input
    - Rich text: Formatted text with paragraphs, runs, etc.
    - Date picker: Calendar-based date selection
    - Dropdown list: Select from predefined options
    - Combo box: Select or type custom value
    - Picture: Image placeholder
    """

    def __init__(self, sdt_elm: CT_SdtBlock, parent: t.ProvidesStoryPart) -> None:
        super().__init__(sdt_elm, parent)
        self._sdt_elm = sdt_elm

    def add_paragraph(
        self, text: str = "", style: str | ParagraphStyle | None = None
    ) -> Paragraph:
        """Return paragraph newly added to the end of this content control.

        The paragraph has `text` in a single run if present, and is given
        paragraph style `style`.
        """
        return super().add_paragraph(text, style)

    @property
    def tag(self) -> str | None:
        """Read/write. The tag value of this content control.

        The tag is a user-defined string that can be used to identify content
        controls programmatically. It is not visible in the Word UI.

        Returns None if no tag is set.
        """
        return self._sdt_elm.sdt_tag

    @tag.setter
    def tag(self, value: str | None) -> None:
        self._sdt_elm.sdt_tag = value

    @property
    def alias(self) -> str | None:
        """Read/write. The alias (title) of this content control.

        The alias is displayed in the Word UI as the content control's title
        when the control is selected. It can be used as a user-friendly label.

        Returns None if no alias is set.
        """
        return self._sdt_elm.alias

    @alias.setter
    def alias(self, value: str | None) -> None:
        self._sdt_elm.alias = value

    @property
    def sdt_type(self) -> str:
        """Read-only. The type of this content control.

        Returns one of: 'text', 'date', 'dropDownList', 'comboBox',
        'docPartObj', 'picture', 'richText'.

        - 'text': Plain text control (single or multi-line)
        - 'date': Date picker control
        - 'dropDownList': Dropdown list (fixed choices)
        - 'comboBox': Combo box (fixed choices or custom input)
        - 'docPartObj': Document part (e.g., table of contents)
        - 'picture': Picture placeholder
        - 'richText': Rich text control (default when no type specified)
        """
        return self._sdt_elm.sdt_type

    @property
    def text(self) -> str:
        """Read-only. The text content of this content control.

        Returns all text from paragraphs in this content control, with
        paragraph breaks represented as newlines.

        For more control over content, use `.paragraphs` to access
        individual paragraphs and their runs.
        """
        return "\n".join(p.text for p in self.paragraphs)


class SdtRunContentControl:
    """Proxy for an inline/run-level structured document tag (content control).

    An inline content control appears within a paragraph and contains runs
    rather than block-level content. It is used for inline form fields.
    """

    def __init__(self, sdt_elm: CT_SdtRun, parent: t.ProvidesStoryPart) -> None:
        self._sdt_elm = sdt_elm
        self._parent = parent

    @property
    def tag(self) -> str | None:
        """Read-only. The tag value of this content control."""
        return self._sdt_elm.sdt_tag

    @property
    def alias(self) -> str | None:
        """Read-only. The alias (title) of this content control."""
        return self._sdt_elm.alias

    @property
    def sdt_type(self) -> str:
        """Read-only. The type of this content control."""
        return self._sdt_elm.sdt_type

    @property
    def text(self) -> str:
        """Read-only. The text content of this inline content control."""
        return self._sdt_elm.text
