"""Custom element classes for shape-related elements like `<w:inline>`."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.simpletypes import (
    ST_Coordinate,
    ST_DrawingElementId,
    ST_PositiveCoordinate,
    ST_RelationshipId,
    XsdString,
    XsdToken,
)
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    OptionalAttribute,
    RequiredAttribute,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.shared import Length


class CT_Anchor(BaseOxmlElement):
    """`<wp:anchor>` element, container for a "floating" shape.

    Floating shapes are positioned relative to page, margin, column, or paragraph
    and can have text wrap around them.

    Key attributes:
        distT/distB/distL/distR: Distance from text in EMUs
        behindDoc: True if shape is behind document text
        relativeHeight: Z-order relative to other shapes
        locked: True if position is locked
        layoutInCell: True if positioned relative to table cell
        allowOverlap: True if can overlap other shapes

    Key children:
        wp:positionH: Horizontal position specification
        wp:positionV: Vertical position specification
        wp:extent: Size (cx, cy) in EMUs
        wp:wrap*: Text wrapping specification (Square, Tight, Through, TopAndBottom, None)
        wp:docPr: Drawing properties (id, name, description)
        a:graphic: The actual graphic content
    """

    extent: CT_PositiveSize2D = ZeroOrOne("wp:extent")  # pyright: ignore[reportAssignmentType]
    docPr: CT_NonVisualDrawingProps = ZeroOrOne("wp:docPr")  # pyright: ignore[reportAssignmentType]
    graphic: CT_GraphicalObject = ZeroOrOne("a:graphic")  # pyright: ignore[reportAssignmentType]

    # Position attributes
    distT: int | None = OptionalAttribute("distT", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    distB: int | None = OptionalAttribute("distB", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    distL: int | None = OptionalAttribute("distL", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    distR: int | None = OptionalAttribute("distR", ST_Coordinate)  # pyright: ignore[reportAssignmentType]
    behindDoc: bool | None = OptionalAttribute("behindDoc", XsdString)  # pyright: ignore[reportAssignmentType]
    relativeHeight: int | None = OptionalAttribute("relativeHeight", ST_Coordinate)  # pyright: ignore[reportAssignmentType]

    @property
    def is_behind_text(self) -> bool:
        """True if this shape is positioned behind document text."""
        return self.behindDoc == "1" or self.behindDoc == "true"

    @classmethod
    def new(
        cls,
        cx: Length,
        cy: Length,
        shape_id: int,
        pic: CT_Picture,
        pos_x: Length,
        pos_y: Length,
        behind_doc: bool = False,
        relative_height: int = 251658240,
        wrap_type: str = "square",
        h_relative_from: str = "column",
        v_relative_from: str = "paragraph",
    ) -> CT_Anchor:
        """Create a new `<wp:anchor>` element for a floating picture.

        Args:
            cx: Width in EMUs
            cy: Height in EMUs
            shape_id: Unique shape identifier
            pic: The CT_Picture element containing the image
            pos_x: Horizontal position offset in EMUs
            pos_y: Vertical position offset in EMUs
            behind_doc: If True, shape is behind document text
            relative_height: Z-order (higher = on top)
            wrap_type: Wrapping style ('none', 'square', 'tight', 'through', 'topAndBottom')
            h_relative_from: Horizontal position relative to ('column', 'page', 'margin', etc.)
            v_relative_from: Vertical position relative to ('paragraph', 'page', 'margin', etc.)

        Returns:
            A new CT_Anchor element.
        """
        anchor = cast(CT_Anchor, parse_xml(cls._anchor_xml(
            cx, cy, shape_id, pos_x, pos_y, behind_doc, relative_height,
            wrap_type, h_relative_from, v_relative_from
        )))
        anchor.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def new_pic_anchor(
        cls,
        shape_id: int,
        rId: str,
        filename: str,
        cx: Length,
        cy: Length,
        pos_x: Length,
        pos_y: Length,
        behind_doc: bool = False,
        wrap_type: str = "square",
        h_relative_from: str = "column",
        v_relative_from: str = "paragraph",
    ) -> CT_Anchor:
        """Create `wp:anchor` element containing a `pic:pic` element.

        Args:
            shape_id: Unique shape identifier
            rId: Relationship ID for the image
            filename: Image filename
            cx: Width in EMUs
            cy: Height in EMUs
            pos_x: Horizontal position offset in EMUs
            pos_y: Vertical position offset in EMUs
            behind_doc: If True, shape is behind document text
            wrap_type: Wrapping style ('none', 'square', 'tight', 'through', 'topAndBottom')
            h_relative_from: Horizontal position relative to
            v_relative_from: Vertical position relative to

        Returns:
            A new CT_Anchor element with picture content.
        """
        pic_id = 0
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        anchor = cls.new(
            cx, cy, shape_id, pic, pos_x, pos_y,
            behind_doc=behind_doc,
            wrap_type=wrap_type,
            h_relative_from=h_relative_from,
            v_relative_from=v_relative_from,
        )
        return anchor

    @classmethod
    def _anchor_xml(
        cls,
        cx: Length,
        cy: Length,
        shape_id: int,
        pos_x: Length,
        pos_y: Length,
        behind_doc: bool,
        relative_height: int,
        wrap_type: str,
        h_relative_from: str,
        v_relative_from: str,
    ) -> str:
        """Generate the XML template for an anchor element."""
        behind_doc_val = "1" if behind_doc else "0"

        # Select wrap element based on wrap_type
        wrap_elements = {
            "none": "<wp:wrapNone/>",
            "square": '<wp:wrapSquare wrapText="bothSides"/>',
            "tight": '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>',
            "through": '<wp:wrapThrough wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapThrough>',
            "topAndBottom": "<wp:wrapTopAndBottom/>",
        }
        wrap_xml = wrap_elements.get(wrap_type, wrap_elements["square"])

        return (
            '<wp:anchor %s distT="0" distB="0" distL="114300" distR="114300" '
            'simplePos="0" relativeHeight="%d" behindDoc="%s" locked="0" '
            'layoutInCell="1" allowOverlap="1">\n'
            '  <wp:simplePos x="0" y="0"/>\n'
            '  <wp:positionH relativeFrom="%s">\n'
            '    <wp:posOffset>%d</wp:posOffset>\n'
            '  </wp:positionH>\n'
            '  <wp:positionV relativeFrom="%s">\n'
            '    <wp:posOffset>%d</wp:posOffset>\n'
            '  </wp:positionV>\n'
            '  <wp:extent cx="%d" cy="%d"/>\n'
            "  <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>\n"
            "  %s\n"
            '  <wp:docPr id="%d" name="Picture %d"/>\n'
            "  <wp:cNvGraphicFramePr>\n"
            '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
            "  </wp:cNvGraphicFramePr>\n"
            "  <a:graphic>\n"
            '    <a:graphicData uri="URI not set"/>\n'
            "  </a:graphic>\n"
            "</wp:anchor>"
            % (
                nsdecls("wp", "a", "pic", "r"),
                relative_height,
                behind_doc_val,
                h_relative_from,
                int(pos_x),
                v_relative_from,
                int(pos_y),
                int(cx),
                int(cy),
                wrap_xml,
                shape_id,
                shape_id,
            )
        )


class CT_Blip(BaseOxmlElement):
    """``<a:blip>`` element, specifies image source and adjustments such as alpha and
    tint."""

    embed: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:embed", ST_RelationshipId
    )
    link: str | None = OptionalAttribute(  # pyright: ignore[reportAssignmentType]
        "r:link", ST_RelationshipId
    )


class CT_BlipFillProperties(BaseOxmlElement):
    """``<pic:blipFill>`` element, specifies picture properties."""

    blip: CT_Blip = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "a:blip", successors=("a:srcRect", "a:tile", "a:stretch")
    )


class CT_GraphicalObject(BaseOxmlElement):
    """``<a:graphic>`` element, container for a DrawingML object."""

    graphicData: CT_GraphicalObjectData = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphicData"
    )


class CT_GraphicalObjectData(BaseOxmlElement):
    """``<a:graphicData>`` element, container for the XML of a DrawingML object."""

    pic: CT_Picture = ZeroOrOne("pic:pic")  # pyright: ignore[reportAssignmentType]
    uri: str = RequiredAttribute("uri", XsdToken)  # pyright: ignore[reportAssignmentType]


class CT_Inline(BaseOxmlElement):
    """`<wp:inline>` element, container for an inline shape."""

    extent: CT_PositiveSize2D = OneAndOnlyOne("wp:extent")  # pyright: ignore[reportAssignmentType]
    docPr: CT_NonVisualDrawingProps = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "wp:docPr"
    )
    graphic: CT_GraphicalObject = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "a:graphic"
    )

    @classmethod
    def new(cls, cx: Length, cy: Length, shape_id: int, pic: CT_Picture) -> CT_Inline:
        """Return a new ``<wp:inline>`` element populated with the values passed as
        parameters."""
        inline = cast(CT_Inline, parse_xml(cls._inline_xml()))
        inline.extent.cx = cx
        inline.extent.cy = cy
        inline.docPr.id = shape_id
        inline.docPr.name = "Picture %d" % shape_id
        inline.graphic.graphicData.uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
        inline.graphic.graphicData._insert_pic(pic)
        return inline

    @classmethod
    def new_pic_inline(
        cls, shape_id: int, rId: str, filename: str, cx: Length, cy: Length
    ) -> CT_Inline:
        """Create `wp:inline` element containing a `pic:pic` element.

        The contents of the `pic:pic` element is taken from the argument values.
        """
        pic_id = 0  # Word doesn't seem to use this, but does not omit it
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        inline = cls.new(cx, cy, shape_id, pic)
        return inline

    @classmethod
    def _inline_xml(cls):
        return (
            "<wp:inline %s>\n"
            '  <wp:extent cx="914400" cy="914400"/>\n'
            '  <wp:docPr id="666" name="unnamed"/>\n'
            "  <wp:cNvGraphicFramePr>\n"
            '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
            "  </wp:cNvGraphicFramePr>\n"
            "  <a:graphic>\n"
            '    <a:graphicData uri="URI not set"/>\n'
            "  </a:graphic>\n"
            "</wp:inline>" % nsdecls("wp", "a", "pic", "r")
        )


class CT_NonVisualDrawingProps(BaseOxmlElement):
    """Used for ``<wp:docPr>`` element, and perhaps others.

    Specifies the id and name of a DrawingML drawing.
    """

    id = RequiredAttribute("id", ST_DrawingElementId)
    name = RequiredAttribute("name", XsdString)
    descr: str | None = OptionalAttribute("descr", XsdString)  # pyright: ignore[reportAssignmentType]


class CT_NonVisualPictureProperties(BaseOxmlElement):
    """``<pic:cNvPicPr>`` element, specifies picture locking and resize behaviors."""


class CT_Picture(BaseOxmlElement):
    """``<pic:pic>`` element, a DrawingML picture."""

    nvPicPr: CT_PictureNonVisual = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "pic:nvPicPr"
    )
    blipFill: CT_BlipFillProperties = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "pic:blipFill"
    )
    spPr: CT_ShapeProperties = OneAndOnlyOne("pic:spPr")  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, pic_id: int, filename: str, rId: str, cx: Length, cy: Length) -> CT_Picture:
        """A new minimum viable `<pic:pic>` (picture) element."""
        pic = parse_xml(cls._pic_xml())
        pic.nvPicPr.cNvPr.id = pic_id
        pic.nvPicPr.cNvPr.name = filename
        pic.blipFill.blip.embed = rId
        pic.spPr.cx = cx
        pic.spPr.cy = cy
        return pic

    @classmethod
    def _pic_xml(cls):
        return (
            "<pic:pic %s>\n"
            "  <pic:nvPicPr>\n"
            '    <pic:cNvPr id="666" name="unnamed"/>\n'
            "    <pic:cNvPicPr/>\n"
            "  </pic:nvPicPr>\n"
            "  <pic:blipFill>\n"
            "    <a:blip/>\n"
            "    <a:stretch>\n"
            "      <a:fillRect/>\n"
            "    </a:stretch>\n"
            "  </pic:blipFill>\n"
            "  <pic:spPr>\n"
            "    <a:xfrm>\n"
            '      <a:off x="0" y="0"/>\n'
            '      <a:ext cx="914400" cy="914400"/>\n'
            "    </a:xfrm>\n"
            '    <a:prstGeom prst="rect"/>\n'
            "  </pic:spPr>\n"
            "</pic:pic>" % nsdecls("pic", "a", "r")
        )


class CT_PictureNonVisual(BaseOxmlElement):
    """``<pic:nvPicPr>`` element, non-visual picture properties."""

    cNvPr = OneAndOnlyOne("pic:cNvPr")


class CT_Point2D(BaseOxmlElement):
    """Used for ``<a:off>`` element, and perhaps others.

    Specifies an x, y coordinate (point).
    """

    x = RequiredAttribute("x", ST_Coordinate)
    y = RequiredAttribute("y", ST_Coordinate)


class CT_PositiveSize2D(BaseOxmlElement):
    """Used for ``<wp:extent>`` element, and perhaps others later.

    Specifies the size of a DrawingML drawing.
    """

    cx: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cx", ST_PositiveCoordinate
    )
    cy: Length = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "cy", ST_PositiveCoordinate
    )


class CT_PresetGeometry2D(BaseOxmlElement):
    """``<a:prstGeom>`` element, specifies an preset autoshape geometry, such as
    ``rect``."""


class CT_RelativeRect(BaseOxmlElement):
    """``<a:fillRect>`` element, specifying picture should fill containing rectangle
    shape."""


class CT_ShapeProperties(BaseOxmlElement):
    """``<pic:spPr>`` element, specifies size and shape of picture container."""

    xfrm = ZeroOrOne(
        "a:xfrm",
        successors=(
            "a:custGeom",
            "a:prstGeom",
            "a:ln",
            "a:effectLst",
            "a:effectDag",
            "a:scene3d",
            "a:sp3d",
            "a:extLst",
        ),
    )

    @property
    def cx(self):
        """Shape width as an instance of Emu, or None if not present."""
        xfrm = self.xfrm
        if xfrm is None:
            return None
        return xfrm.cx

    @cx.setter
    def cx(self, value):
        xfrm = self.get_or_add_xfrm()
        xfrm.cx = value

    @property
    def cy(self):
        """Shape height as an instance of Emu, or None if not present."""
        xfrm = self.xfrm
        if xfrm is None:
            return None
        return xfrm.cy

    @cy.setter
    def cy(self, value):
        xfrm = self.get_or_add_xfrm()
        xfrm.cy = value


class CT_StretchInfoProperties(BaseOxmlElement):
    """``<a:stretch>`` element, specifies how picture should fill its containing
    shape."""


class CT_Transform2D(BaseOxmlElement):
    """``<a:xfrm>`` element, specifies size and shape of picture container."""

    off = ZeroOrOne("a:off", successors=("a:ext",))
    ext = ZeroOrOne("a:ext", successors=())

    @property
    def cx(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cx

    @cx.setter
    def cx(self, value):
        ext = self.get_or_add_ext()
        ext.cx = value

    @property
    def cy(self):
        ext = self.ext
        if ext is None:
            return None
        return ext.cy

    @cy.setter
    def cy(self, value):
        ext = self.get_or_add_ext()
        ext.cy = value
