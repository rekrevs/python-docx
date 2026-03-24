"""|NumberingPart| and closely related objects."""

from __future__ import annotations

from typing import TYPE_CHECKING

from ..opc.packuri import PackURI
from ..opc.part import XmlPart
from ..shared import lazyproperty

if TYPE_CHECKING:
    from ..oxml.numbering import CT_AbstractNum, CT_Num, CT_Numbering


class NumberingPart(XmlPart):
    """Proxy for the numbering.xml part containing numbering definitions for a document
    or glossary."""

    @classmethod
    def new(cls) -> NumberingPart:
        """Return a newly created NumberingPart with minimal numbering definitions.

        The new part contains the root ``<w:numbering>`` element with basic abstract
        numbering definitions for bullet and numbered lists.
        """
        from ..opc.constants import CONTENT_TYPE as CT
        from ..oxml.ns import nsdecls
        from ..oxml.parser import parse_xml

        partname = PackURI("/word/numbering.xml")

        # Create minimal numbering.xml with basic bullet and decimal abstract nums
        numbering_xml = (
            f'<w:numbering {nsdecls("w", "r")}>'
            f'  <w:abstractNum w:abstractNumId="0">'
            f'    <w:nsid w:val="00000001"/>'
            f'    <w:multiLevelType w:val="singleLevel"/>'
            f'    <w:lvl w:ilvl="0">'
            f'      <w:start w:val="1"/>'
            f'      <w:numFmt w:val="bullet"/>'
            f'      <w:lvlText w:val="&#61623;"/>'
            f'      <w:lvlJc w:val="left"/>'
            f'      <w:pPr>'
            f'        <w:ind w:left="720" w:hanging="360"/>'
            f'      </w:pPr>'
            f'      <w:rPr>'
            f'        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>'
            f'      </w:rPr>'
            f'    </w:lvl>'
            f'  </w:abstractNum>'
            f'  <w:abstractNum w:abstractNumId="1">'
            f'    <w:nsid w:val="00000002"/>'
            f'    <w:multiLevelType w:val="singleLevel"/>'
            f'    <w:lvl w:ilvl="0">'
            f'      <w:start w:val="1"/>'
            f'      <w:numFmt w:val="decimal"/>'
            f'      <w:lvlText w:val="%1."/>'
            f'      <w:lvlJc w:val="left"/>'
            f'      <w:pPr>'
            f'        <w:ind w:left="720" w:hanging="360"/>'
            f'      </w:pPr>'
            f'    </w:lvl>'
            f'  </w:abstractNum>'
            f'  <w:num w:numId="1">'
            f'    <w:abstractNumId w:val="0"/>'
            f'  </w:num>'
            f'  <w:num w:numId="2">'
            f'    <w:abstractNumId w:val="1"/>'
            f'  </w:num>'
            f'</w:numbering>'
        )

        element = parse_xml(numbering_xml)
        content_type = CT.WML_NUMBERING

        return cls(partname, content_type, element, package=None)

    @property
    def numbering_elm(self) -> CT_Numbering:
        """The `w:numbering` root element of this numbering part."""
        return self._element  # type: ignore[return-value]

    def add_abstract_num(
        self,
        multi_level_type: str = "hybridMultilevel",
    ) -> CT_AbstractNum:
        """Add a new abstract numbering definition and return it.

        Args:
            multi_level_type: One of "singleLevel", "multiLevel", "hybridMultilevel".

        Returns:
            The newly created CT_AbstractNum element.
        """
        return self.numbering_elm.add_abstractNum(multi_level_type)

    def add_num(self, abstract_num_id: int) -> CT_Num:
        """Add a new `w:num` element referencing `abstract_num_id` and return it.

        Args:
            abstract_num_id: The abstractNumId of the abstract definition to reference.

        Returns:
            The newly created CT_Num element with an auto-assigned numId.
        """
        return self.numbering_elm.add_num(abstract_num_id)

    @lazyproperty
    def numbering_definitions(self):
        """The |_NumberingDefinitions| instance containing the numbering definitions
        (<w:num> element proxies) for this numbering part."""
        return _NumberingDefinitions(self._element)


class _NumberingDefinitions:
    """Collection of |_NumberingDefinition| instances corresponding to the ``<w:num>``
    elements in a numbering part."""

    def __init__(self, numbering_elm):
        super(_NumberingDefinitions, self).__init__()
        self._numbering = numbering_elm

    def __len__(self):
        return len(self._numbering.num_lst)
