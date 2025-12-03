"""Footnotes and endnotes proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.footnotes import CT_Endnote, CT_Footnote
    from docx.parts.footnotes import EndnotesPart, FootnotesPart
    from docx.text.paragraph import Paragraph


class Footnotes:
    """Collection of footnotes in a document.

    Provides access to the footnotes defined in the document's footnotes part.
    Footnotes with type='separator' or 'continuationSeparator' are filtered out
    as they are used for rendering purposes only.
    """

    def __init__(self, footnotes_part: FootnotesPart | None) -> None:
        self._footnotes_part = footnotes_part

    def __iter__(self) -> Iterator[Footnote]:
        """Iterate over all user-visible footnotes (excludes separators)."""
        if self._footnotes_part is None:
            return
        for fn_elm in self._footnotes_part.footnotes:
            if not fn_elm.is_separator:
                yield Footnote(fn_elm, self._footnotes_part)

    def __len__(self) -> int:
        """Return the number of user-visible footnotes."""
        if self._footnotes_part is None:
            return 0
        return sum(1 for fn in self._footnotes_part.footnotes if not fn.is_separator)

    def __getitem__(self, footnote_id: int) -> Footnote | None:
        """Return the footnote with the specified id, or None if not found.

        Note: Footnote IDs start at 1 for user-visible footnotes. IDs -1 and 0
        are reserved for separator footnotes.
        """
        if self._footnotes_part is None:
            return None
        fn_elm = self._footnotes_part.footnote_by_id(footnote_id)
        if fn_elm is None:
            return None
        return Footnote(fn_elm, self._footnotes_part)


class Footnote(BlockItemContainer):
    """Proxy for a single footnote.

    A footnote contains paragraphs and tables, similar to the document body.
    """

    def __init__(self, footnote_elm: CT_Footnote, part: FootnotesPart) -> None:
        super().__init__(footnote_elm, part)
        self._footnote_elm = footnote_elm

    @property
    def footnote_id(self) -> int:
        """The id of this footnote.

        This id is referenced by `<w:footnoteReference>` elements in the document body.
        """
        return self._footnote_elm.footnote_id

    @property
    def text(self) -> str:
        """The concatenated text of all paragraphs in this footnote."""
        return "".join(p.text or "" for p in self.paragraphs)


class Endnotes:
    """Collection of endnotes in a document.

    Provides access to the endnotes defined in the document's endnotes part.
    Endnotes with type='separator' or 'continuationSeparator' are filtered out
    as they are used for rendering purposes only.
    """

    def __init__(self, endnotes_part: EndnotesPart | None) -> None:
        self._endnotes_part = endnotes_part

    def __iter__(self) -> Iterator[Endnote]:
        """Iterate over all user-visible endnotes (excludes separators)."""
        if self._endnotes_part is None:
            return
        for en_elm in self._endnotes_part.endnotes:
            if not en_elm.is_separator:
                yield Endnote(en_elm, self._endnotes_part)

    def __len__(self) -> int:
        """Return the number of user-visible endnotes."""
        if self._endnotes_part is None:
            return 0
        return sum(1 for en in self._endnotes_part.endnotes if not en.is_separator)

    def __getitem__(self, endnote_id: int) -> Endnote | None:
        """Return the endnote with the specified id, or None if not found.

        Note: Endnote IDs start at 1 for user-visible endnotes. IDs -1 and 0
        are reserved for separator endnotes.
        """
        if self._endnotes_part is None:
            return None
        en_elm = self._endnotes_part.endnote_by_id(endnote_id)
        if en_elm is None:
            return None
        return Endnote(en_elm, self._endnotes_part)


class Endnote(BlockItemContainer):
    """Proxy for a single endnote.

    An endnote contains paragraphs and tables, similar to the document body.
    """

    def __init__(self, endnote_elm: CT_Endnote, part: EndnotesPart) -> None:
        super().__init__(endnote_elm, part)
        self._endnote_elm = endnote_elm

    @property
    def endnote_id(self) -> int:
        """The id of this endnote.

        This id is referenced by `<w:endnoteReference>` elements in the document body.
        """
        return self._endnote_elm.endnote_id

    @property
    def text(self) -> str:
        """The concatenated text of all paragraphs in this endnote."""
        return "".join(p.text or "" for p in self.paragraphs)
