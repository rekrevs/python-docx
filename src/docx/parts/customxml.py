"""Provides CustomXmlPart and related objects."""

# pyright: reportPrivateUsage=false

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Any, List

from lxml import etree

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part, XmlPart

if TYPE_CHECKING:
    from docx.oxml.customxml import CT_DatastoreItem
    from docx.package import Package


class CustomXmlPart(Part):
    """A custom XML data part containing arbitrary XML content.

    Custom XML parts store user-defined XML data that can be bound to content
    controls in the document. Each custom XML part has an associated properties
    part (CustomXmlPropertiesPart) containing the itemID (GUID) and schema refs.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        blob: bytes | None = None,
        package: Package | None = None,
    ):
        super().__init__(partname, content_type, blob, package)
        self._xml_element: etree._Element | None = None

    @property
    def element(self) -> etree._Element | None:
        """Return the root XML element of this custom XML part."""
        if self._xml_element is None and self._blob:
            self._xml_element = etree.fromstring(self._blob)
        return self._xml_element

    @element.setter
    def element(self, value: etree._Element) -> None:
        """Set the root XML element."""
        self._xml_element = value

    @property
    def blob(self) -> bytes:
        """Return serialized XML content."""
        if self._xml_element is not None:
            return etree.tostring(
                self._xml_element,
                encoding="UTF-8",
                xml_declaration=True,
                standalone=True,
            )
        return self._blob or b""

    def xpath(self, xpath_expr: str, namespaces: dict[str, str] | None = None) -> List[Any]:
        """Execute an XPath query on the custom XML content.

        Args:
            xpath_expr: XPath expression to evaluate.
            namespaces: Optional namespace prefix mappings.

        Returns:
            List of matching elements or attribute values.
        """
        element = self.element
        if element is None:
            return []
        return element.xpath(xpath_expr, namespaces=namespaces or {})

    @classmethod
    def new(
        cls,
        package: Package,
        xml_element: etree._Element,
        item_num: int = 1,
    ) -> tuple[CustomXmlPart, CustomXmlPropertiesPart]:
        """Create a new custom XML part with the given XML content.

        Args:
            package: The package this part belongs to.
            xml_element: The root XML element for the custom XML content.
            item_num: The item number (1-based) for the part name.

        Returns:
            A tuple of (CustomXmlPart, CustomXmlPropertiesPart).
        """
        # Create the custom XML part
        partname = PackURI(f"/customXml/item{item_num}.xml")
        content_type = CT.XML
        xml_blob = etree.tostring(
            xml_element,
            encoding="UTF-8",
            xml_declaration=True,
            standalone=True,
        )
        custom_xml_part = cls(partname, content_type, xml_blob, package)

        # Create the properties part with a new GUID
        item_id = "{" + str(uuid.uuid4()).upper() + "}"
        props_part = CustomXmlPropertiesPart.new(package, item_id, item_num)

        # Create relationship from custom XML to its properties
        custom_xml_part.relate_to(props_part, RT.CUSTOM_XML_PROPS)

        return custom_xml_part, props_part


class CustomXmlPropertiesPart(XmlPart):
    """Properties part for a custom XML part.

    Contains the datastoreItem element with itemID (GUID) and schema references.
    """

    @property
    def datastore_item(self) -> CT_DatastoreItem:
        """Return the root ``<ds:datastoreItem>`` element."""
        return self._element  # type: ignore[return-value]

    @property
    def item_id(self) -> str | None:
        """Return the itemID (GUID) for this custom XML part."""
        return self.datastore_item.item_id

    @property
    def schema_uris(self) -> list[str]:
        """Return list of schema URIs referenced by this custom XML part."""
        return self.datastore_item.schema_uris

    @classmethod
    def new(cls, package: Package, item_id: str, item_num: int = 1) -> CustomXmlPropertiesPart:
        """Create a new custom XML properties part.

        Args:
            package: The package this part belongs to.
            item_id: The GUID for the custom XML part (including braces).
            item_num: The item number (1-based) for the part name.

        Returns:
            A new CustomXmlPropertiesPart.
        """
        from docx.oxml.customxml import CT_DatastoreItem

        partname = PackURI(f"/customXml/itemProps{item_num}.xml")
        content_type = CT.OFC_CUSTOM_XML_PROPERTIES
        element = CT_DatastoreItem.new(item_id)
        return cls(partname, content_type, element, package)
