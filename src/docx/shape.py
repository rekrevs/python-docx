"""Objects related to shapes.

A shape is a visual object that appears on the drawing layer of a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.enum.shape import WD_INLINE_SHAPE
from docx.oxml.ns import nsmap
from docx.shared import Parented

if TYPE_CHECKING:
    from docx.oxml.document import CT_Body
    from docx.oxml.shape import CT_Anchor, CT_Inline
    from docx.parts.story import StoryPart
    from docx.shared import Length


class InlineShapes(Parented):
    """Sequence of |InlineShape| instances, supporting len(), iteration, and indexed access."""

    def __init__(self, body_elm: CT_Body, parent: StoryPart):
        super(InlineShapes, self).__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx: int):
        """Provide indexed access, e.g. 'inline_shapes[idx]'."""
        try:
            inline = self._inline_lst[idx]
        except IndexError:
            msg = "inline shape index [%d] out of range" % idx
            raise IndexError(msg)

        return InlineShape(inline)

    def __iter__(self):
        return (InlineShape(inline) for inline in self._inline_lst)

    def __len__(self):
        return len(self._inline_lst)

    @property
    def _inline_lst(self):
        body = self._body
        xpath = "//w:p/w:r/w:drawing/wp:inline"
        return body.xpath(xpath)


class InlineShape:
    """Proxy for an ``<wp:inline>`` element, representing the container for an inline
    graphical object."""

    def __init__(self, inline: CT_Inline):
        super(InlineShape, self).__init__()
        self._inline = inline

    @property
    def height(self) -> Length:
        """Read/write.

        The display height of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cy

    @height.setter
    def height(self, cy: Length):
        self._inline.extent.cy = cy
        self._inline.graphic.graphicData.pic.spPr.cy = cy

    @property
    def type(self):
        """The type of this inline shape as a member of
        ``docx.enum.shape.WD_INLINE_SHAPE``, e.g. ``LINKED_PICTURE``.

        Read-only.
        """
        graphicData = self._inline.graphic.graphicData
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            blip = graphicData.pic.blipFill.blip
            if blip.link is not None:
                return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED

    @property
    def width(self):
        """Read/write.

        The display width of this inline shape as an |Emu| instance.
        """
        return self._inline.extent.cx

    @width.setter
    def width(self, cx: Length):
        self._inline.extent.cx = cx
        self._inline.graphic.graphicData.pic.spPr.cx = cx


class FloatingShapes(Parented):
    """Sequence of |FloatingShape| instances for anchored/floating shapes.

    Floating shapes (also called anchored shapes) are positioned independently
    of the text flow and can have text wrap around them.

    Example::

        # Iterate over floating shapes
        for shape in document.floating_shapes:
            print(f"Shape: {shape.name}, size: {shape.width}x{shape.height}")

        # Access by index
        first_shape = document.floating_shapes[0]
    """

    def __init__(self, body_elm: CT_Body, parent: StoryPart):
        super().__init__(parent)
        self._body = body_elm

    def __getitem__(self, idx: int) -> FloatingShape:
        """Provide indexed access, e.g. 'floating_shapes[idx]'."""
        try:
            anchor = self._anchor_lst[idx]
        except IndexError:
            raise IndexError(f"floating shape index [{idx}] out of range")
        return FloatingShape(anchor, self.part)

    def __iter__(self) -> Iterator[FloatingShape]:
        """Iterate over all floating shapes."""
        return (FloatingShape(anchor, self.part) for anchor in self._anchor_lst)

    def __len__(self) -> int:
        """Return the number of floating shapes."""
        return len(self._anchor_lst)

    @property
    def _anchor_lst(self):
        """List of all wp:anchor elements in the document body."""
        return self._body.xpath("//w:p/w:r/w:drawing/wp:anchor")


class FloatingShape(Parented):
    """Proxy for a ``<wp:anchor>`` element, representing a floating/anchored shape.

    Floating shapes are positioned relative to page elements (page, margin,
    column, paragraph, line, character) rather than flowing inline with text.
    Text can wrap around floating shapes in various ways.

    Attributes:
        width: Display width in EMUs
        height: Display height in EMUs
        name: The name of the shape
        description: Alt text description
        is_behind_text: True if shape is behind document text
        type: The shape type (PICTURE, CHART, etc.)
    """

    def __init__(self, anchor: CT_Anchor, parent: StoryPart):
        super().__init__(parent)
        self._anchor = anchor

    @property
    def height(self) -> Length | None:
        """The display height of this floating shape in EMUs, or None if not set."""
        extent = self._anchor.extent
        return extent.cy if extent is not None else None

    @property
    def width(self) -> Length | None:
        """The display width of this floating shape in EMUs, or None if not set."""
        extent = self._anchor.extent
        return extent.cx if extent is not None else None

    @property
    def name(self) -> str:
        """The name of this shape, or empty string if not set."""
        docPr = self._anchor.docPr
        return docPr.name if docPr is not None else ""

    @property
    def description(self) -> str:
        """The description (alt text) of this shape, or empty string if not set."""
        docPr = self._anchor.docPr
        return docPr.descr if docPr is not None else ""

    @property
    def is_behind_text(self) -> bool:
        """True if this shape is positioned behind document text."""
        return self._anchor.is_behind_text

    @property
    def type(self):
        """The type of this floating shape as a member of WD_INLINE_SHAPE enum.

        Note: Uses the same enum as inline shapes since the content types are the same.
        """
        graphic = self._anchor.graphic
        if graphic is None:
            return WD_INLINE_SHAPE.NOT_IMPLEMENTED
        graphicData = graphic.graphicData
        if graphicData is None:
            return WD_INLINE_SHAPE.NOT_IMPLEMENTED
        uri = graphicData.uri
        if uri == nsmap["pic"]:
            pic = graphicData.pic
            if pic is not None:
                blip = pic.blipFill.blip
                if blip is not None and blip.link is not None:
                    return WD_INLINE_SHAPE.LINKED_PICTURE
            return WD_INLINE_SHAPE.PICTURE
        if uri == nsmap["c"]:
            return WD_INLINE_SHAPE.CHART
        if uri == nsmap["dgm"]:
            return WD_INLINE_SHAPE.SMART_ART
        return WD_INLINE_SHAPE.NOT_IMPLEMENTED
