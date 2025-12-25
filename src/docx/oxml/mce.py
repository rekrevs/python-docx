"""Custom element classes for Markup Compatibility and Extensibility (MCE).

MCE is defined in ECMA-376 Part 3 / ISO 29500-3 and provides mechanisms for
forward compatibility in OOXML documents.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import nsdecls
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.shared import Length


class CT_AlternateContent(BaseOxmlElement):
    """`mc:AlternateContent` element.

    Container for version-specific markup alternatives. Contains one or more
    Choice elements followed by an optional Fallback element.
    """

    @property
    def choice(self) -> CT_Choice | None:
        """First `mc:Choice` child element, or None if not present."""
        choices = self.xpath("mc:Choice")
        return choices[0] if choices else None

    @property
    def choices(self) -> List[CT_Choice]:
        """All `mc:Choice` child elements."""
        return self.xpath("mc:Choice")

    @property
    def fallback(self) -> CT_Fallback | None:
        """`mc:Fallback` child element, or None if not present."""
        fallbacks = self.xpath("mc:Fallback")
        return fallbacks[0] if fallbacks else None

    @property
    def txbxContent_choice(self) -> CT_TxbxContent | None:
        """The w:txbxContent element from mc:Choice, if this is a text box."""
        matches = self.xpath("mc:Choice//wps:txbx/w:txbxContent")
        return matches[0] if matches else None

    @property
    def txbxContent_fallback(self) -> CT_TxbxContent | None:
        """The w:txbxContent element from mc:Fallback, if this is a text box."""
        matches = self.xpath("mc:Fallback//v:textbox/w:txbxContent")
        return matches[0] if matches else None

    @property
    def is_textbox(self) -> bool:
        """True if this AlternateContent contains a text box."""
        return self.txbxContent_choice is not None or self.txbxContent_fallback is not None

    @classmethod
    def new_textbox(
        cls,
        shape_id: int,
        cx: Length,
        cy: Length,
        pos_x: Length = 0,  # type: ignore[assignment]
        pos_y: Length = 0,  # type: ignore[assignment]
        wrap_type: str = "square",
    ) -> CT_AlternateContent:
        """Create a new text box in an mc:AlternateContent wrapper.

        Args:
            shape_id: Unique shape identifier
            cx: Width in EMUs
            cy: Height in EMUs
            pos_x: Horizontal position in EMUs (default 0)
            pos_y: Vertical position in EMUs (default 0)
            wrap_type: Text wrapping style ('none', 'square', 'topAndBottom')

        Returns:
            A new CT_AlternateContent element containing the text box.
        """
        wrap_xml_map = {
            "none": "<wp:wrapNone/>",
            "square": '<wp:wrapSquare wrapText="bothSides"/>',
            "topAndBottom": "<wp:wrapTopAndBottom/>",
        }
        wrap_xml = wrap_xml_map.get(wrap_type, wrap_xml_map["square"])

        xml = (
            '<mc:AlternateContent %s>\n'
            '  <mc:Choice Requires="wps">\n'
            '    <w:drawing>\n'
            '      <wp:anchor distT="0" distB="0" distL="114300" distR="114300" '
            '                 simplePos="0" relativeHeight="251659264" behindDoc="0" '
            '                 locked="0" layoutInCell="1" allowOverlap="1">\n'
            '        <wp:simplePos x="0" y="0"/>\n'
            '        <wp:positionH relativeFrom="column">\n'
            '          <wp:posOffset>%d</wp:posOffset>\n'
            '        </wp:positionH>\n'
            '        <wp:positionV relativeFrom="paragraph">\n'
            '          <wp:posOffset>%d</wp:posOffset>\n'
            '        </wp:positionV>\n'
            '        <wp:extent cx="%d" cy="%d"/>\n'
            '        <wp:effectExtent l="0" t="0" r="0" b="0"/>\n'
            '        %s\n'
            '        <wp:docPr id="%d" name="Text Box %d"/>\n'
            '        <wp:cNvGraphicFramePr>\n'
            '          <a:graphicFrameLocks/>\n'
            '        </wp:cNvGraphicFramePr>\n'
            '        <a:graphic>\n'
            '          <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">\n'
            '            <wps:wsp>\n'
            '              <wps:cNvSpPr txBox="1"/>\n'
            '              <wps:spPr>\n'
            '                <a:xfrm>\n'
            '                  <a:off x="0" y="0"/>\n'
            '                  <a:ext cx="%d" cy="%d"/>\n'
            '                </a:xfrm>\n'
            '                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>\n'
            '                <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>\n'
            '                <a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>\n'
            '              </wps:spPr>\n'
            '              <wps:txbx>\n'
            '                <w:txbxContent>\n'
            '                  <w:p/>\n'
            '                </w:txbxContent>\n'
            '              </wps:txbx>\n'
            '              <wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" '
            '                          horzOverflow="overflow" vert="horz" wrap="square" '
            '                          lIns="91440" tIns="45720" rIns="91440" bIns="45720" '
            '                          anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1">\n'
            '                <a:noAutofit/>\n'
            '              </wps:bodyPr>\n'
            '            </wps:wsp>\n'
            '          </a:graphicData>\n'
            '        </a:graphic>\n'
            '      </wp:anchor>\n'
            '    </w:drawing>\n'
            '  </mc:Choice>\n'
            '  <mc:Fallback>\n'
            '    <w:pict>\n'
            '      <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">\n'
            '        <v:stroke joinstyle="miter"/>\n'
            '        <v:path gradientshapeok="t" o:connecttype="rect"/>\n'
            '      </v:shapetype>\n'
            '      <v:shape id="Text Box %d" o:spid="_x0000_s%d" type="#_x0000_t202" '
            '               style="position:absolute;margin-left:%.1fpt;margin-top:%.1fpt;width:%.1fpt;height:%.1fpt;z-index:251659264" '
            '               fillcolor="white" strokeweight=".5pt">\n'
            '        <v:textbox>\n'
            '          <w:txbxContent>\n'
            '            <w:p/>\n'
            '          </w:txbxContent>\n'
            '        </v:textbox>\n'
            '      </v:shape>\n'
            '    </w:pict>\n'
            '  </mc:Fallback>\n'
            '</mc:AlternateContent>'
        ) % (
            nsdecls("mc", "w", "wp", "a", "wps", "v", "o"),
            int(pos_x),
            int(pos_y),
            int(cx),
            int(cy),
            wrap_xml,
            shape_id,
            shape_id,
            int(cx),
            int(cy),
            shape_id,
            shape_id,
            int(pos_x) / 914400 * 72,  # EMU to points
            int(pos_y) / 914400 * 72,
            int(cx) / 914400 * 72,
            int(cy) / 914400 * 72,
        )
        return cast("CT_AlternateContent", parse_xml(xml))


class CT_Choice(BaseOxmlElement):
    """`mc:Choice` element.

    Contains content for applications that understand the namespaces
    specified in the Requires attribute.
    """

    @property
    def requires(self) -> str:
        """The Requires attribute value (space-separated namespace prefixes)."""
        return self.get("Requires", "")

    @property
    def requires_namespaces(self) -> List[str]:
        """List of namespace prefixes required by this Choice."""
        req = self.requires
        return req.split() if req else []


class CT_Fallback(BaseOxmlElement):
    """`mc:Fallback` element.

    Contains fallback content for applications that don't understand
    the namespaces required by any Choice element.
    """
    pass


class CT_TxbxContent(BaseOxmlElement):
    """`w:txbxContent` element.

    Container for text box content. Can contain paragraphs and tables.
    This element appears in both mc:Choice (inside wps:txbx) and
    mc:Fallback (inside v:textbox).
    """

    add_p: Callable[[], CT_P]
    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=())
    tbl = ZeroOrMore("w:tbl", successors=())

    @property
    def p_lst(self) -> List[CT_P]:
        """All `w:p` (paragraph) elements in this text box content."""
        return self.xpath("w:p")

    @property
    def tbl_lst(self) -> List[CT_Tbl]:
        """All `w:tbl` (table) elements in this text box content."""
        return self.xpath("w:tbl")

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` elements in document order."""
        return self.xpath("./w:p | ./w:tbl")
