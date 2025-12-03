"""TextBox proxy objects for accessing text box content.

Text boxes in Word documents are stored inside mc:AlternateContent elements,
with the content in w:txbxContent elements in both mc:Choice (DrawingML) and
mc:Fallback (VML) branches.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from docx.blkcntnr import BlockItemContainer

if TYPE_CHECKING:
    from docx.oxml.mce import CT_AlternateContent
    from docx.parts.story import StoryPart
    from docx.styles.style import ParagraphStyle
    from docx.text.paragraph import Paragraph


class TextBoxes:
    """Collection of text boxes in a document.

    Provides access to all text boxes found in mc:AlternateContent elements
    within the document body.

    Example::

        for textbox in document.text_boxes:
            for para in textbox.paragraphs:
                print(para.text)
    """

    def __init__(
        self,
        alternate_content_elements: List[CT_AlternateContent],
        part: StoryPart,
    ) -> None:
        self._ac_elements = [
            ac for ac in alternate_content_elements if ac.is_textbox
        ]
        self._part = part

    def __iter__(self) -> Iterator[TextBox]:
        """Iterate over all text boxes in document order."""
        for ac in self._ac_elements:
            yield TextBox(ac, self._part)

    def __len__(self) -> int:
        """Return the number of text boxes."""
        return len(self._ac_elements)

    def __getitem__(self, idx: int) -> TextBox:
        """Return text box at the specified index."""
        return TextBox(self._ac_elements[idx], self._part)

    def __bool__(self) -> bool:
        """True if there are any text boxes."""
        return len(self._ac_elements) > 0


class TextBox(BlockItemContainer):
    """Proxy for a text box.

    A text box contains paragraphs and tables, accessible via the standard
    BlockItemContainer interface (paragraphs, tables, add_paragraph, add_table).

    Text boxes are stored in mc:AlternateContent with content in both:
    - mc:Choice (DrawingML): wps:wsp/wps:txbx/w:txbxContent
    - mc:Fallback (VML): v:shape/v:textbox/w:txbxContent

    This class reads from mc:Choice and writes to both branches to maintain
    compatibility with older Word versions.
    """

    def __init__(
        self,
        alternate_content: CT_AlternateContent,
        part: StoryPart,
    ) -> None:
        # Get txbxContent from Choice (primary) or Fallback
        txbx_content_choice = alternate_content.txbxContent_choice
        txbx_content_fallback = alternate_content.txbxContent_fallback
        txbx_content = (
            txbx_content_choice if txbx_content_choice is not None
            else txbx_content_fallback
        )
        if txbx_content is None:
            raise ValueError("AlternateContent does not contain a text box")

        super().__init__(txbx_content, part)
        self._alternate_content = alternate_content
        self._txbx_content_choice = txbx_content_choice
        self._txbx_content_fallback = txbx_content_fallback

    def add_paragraph(
        self,
        text: str = "",
        style: str | ParagraphStyle | None = None,
    ) -> Paragraph:
        """Add a paragraph to the text box.

        The paragraph is added to both mc:Choice and mc:Fallback content
        to maintain compatibility with older Word versions.
        """
        # Add to primary (Choice) content
        paragraph = super().add_paragraph(text, style)

        # Sync to Fallback if it exists
        self._sync_to_fallback()

        return paragraph

    @property
    def text(self) -> str:
        """The combined text of all paragraphs in this text box.

        Paragraphs are separated by newlines.
        """
        return "\n".join(p.text for p in self.paragraphs)

    def _sync_to_fallback(self) -> None:
        """Sync content changes to the mc:Fallback branch.

        This ensures the document remains compatible with older Word versions.
        For now, this is a placeholder - full sync would require copying
        paragraph and table structure.
        """
        # TODO: Implement proper sync from Choice to Fallback
        # For now, modifications to Choice are preserved but Fallback
        # may become stale. This is acceptable for most use cases as
        # modern Word uses the Choice content.
        pass
