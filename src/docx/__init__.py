"""Initialize `docx` package.

Export the `Document` constructor function and establish the mapping of part-type to
the part-classe that implements that type.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Type

from docx.api import Document

if TYPE_CHECKING:
    from docx.opc.part import Part

__version__ = "1.2.0"


__all__ = ["Document"]


# -- register custom Part classes with opc package reader --

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import PartFactory
from docx.opc.parts.coreprops import CorePropertiesPart
from docx.parts.comments import CommentsPart
from docx.parts.customxml import CustomXmlPart, CustomXmlPropertiesPart
from docx.parts.document import DocumentPart
from docx.parts.footnotes import EndnotesPart, FootnotesPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.image import ImagePart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.styles import StylesPart
from docx.parts.theme import ThemePart


def _make_part_class_selector(
    image_part: Type[Part],
    custom_xml_part: Type[Part],
    custom_xml_props_part: Type[Part],
    rt_image: str,
    rt_custom_xml: str,
    rt_custom_xml_props: str,
):
    """Factory to create part_class_selector with captured references."""
    def part_class_selector(content_type: str, reltype: str) -> Type[Part] | None:
        if reltype == rt_image:
            return image_part
        if reltype == rt_custom_xml:
            return custom_xml_part
        if reltype == rt_custom_xml_props:
            return custom_xml_props_part
        return None
    return part_class_selector


part_class_selector = _make_part_class_selector(
    ImagePart,
    CustomXmlPart,
    CustomXmlPropertiesPart,
    RT.IMAGE,
    RT.CUSTOM_XML,
    RT.CUSTOM_XML_PROPS,
)


PartFactory.part_class_selector = part_class_selector
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.WML_COMMENTS] = CommentsPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_ENDNOTES] = EndnotesPart
PartFactory.part_type_for[CT.WML_FOOTER] = FooterPart
PartFactory.part_type_for[CT.WML_FOOTNOTES] = FootnotesPart
PartFactory.part_type_for[CT.WML_HEADER] = HeaderPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart
PartFactory.part_type_for[CT.OFC_THEME] = ThemePart
PartFactory.part_type_for[CT.OFC_CUSTOM_XML_PROPERTIES] = CustomXmlPropertiesPart

del (
    CT,
    CorePropertiesPart,
    CommentsPart,
    CustomXmlPart,
    CustomXmlPropertiesPart,
    DocumentPart,
    EndnotesPart,
    FooterPart,
    FootnotesPart,
    HeaderPart,
    ImagePart,
    NumberingPart,
    PartFactory,
    RT,
    SettingsPart,
    StylesPart,
    ThemePart,
    _make_part_class_selector,
    part_class_selector,
)
