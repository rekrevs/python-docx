"""Proxy objects for custom XML parts.

Custom XML parts allow storing arbitrary XML data within a Word document.
This data can be bound to content controls for data-driven document generation.
"""

# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterator, List

from lxml import etree

if TYPE_CHECKING:
    from docx.parts.customxml import CustomXmlPart, CustomXmlPropertiesPart


class CustomXmlParts:
    """Collection of custom XML parts in a document.

    Provides iteration, indexing, and lookup by item ID (GUID).
    """

    def __init__(
        self,
        custom_xml_parts: List[tuple[CustomXmlPart, CustomXmlPropertiesPart | None]],
    ):
        self._parts = custom_xml_parts

    def __iter__(self) -> Iterator[CustomXml]:
        """Iterate over CustomXml objects."""
        for xml_part, props_part in self._parts:
            yield CustomXml(xml_part, props_part)

    def __len__(self) -> int:
        """Return the number of custom XML parts."""
        return len(self._parts)

    def __getitem__(self, index: int) -> CustomXml:
        """Return CustomXml at `index`."""
        xml_part, props_part = self._parts[index]
        return CustomXml(xml_part, props_part)

    def get_by_item_id(self, item_id: str) -> CustomXml | None:
        """Return the CustomXml with the specified item ID (GUID), or None.

        The item_id should include the curly braces, e.g.
        "{C66AC4F1-97D3-4D10-8D07-51E6CC79AD91}".
        """
        # Normalize to uppercase with braces
        item_id_upper = item_id.upper()
        if not item_id_upper.startswith("{"):
            item_id_upper = "{" + item_id_upper + "}"

        for xml_part, props_part in self._parts:
            if props_part is not None:
                part_id = props_part.item_id
                if part_id and part_id.upper() == item_id_upper:
                    return CustomXml(xml_part, props_part)
        return None

    def get_by_namespace(self, namespace_uri: str) -> CustomXml | None:
        """Return the first CustomXml with the specified root namespace, or None.

        This is useful for finding a specific type of custom XML data when
        you know the namespace URI of the root element.
        """
        for xml_part, props_part in self._parts:
            element = xml_part.element
            if element is not None:
                # Check root element's namespace
                ns = element.nsmap.get(None) or element.nsmap.get(element.prefix)
                if ns == namespace_uri:
                    return CustomXml(xml_part, props_part)
        return None


class CustomXml:
    """Proxy for a single custom XML part and its properties.

    Provides access to the XML content and metadata (item ID, schema URIs).
    """

    def __init__(
        self,
        xml_part: CustomXmlPart,
        properties_part: CustomXmlPropertiesPart | None,
    ):
        self._xml_part = xml_part
        self._properties_part = properties_part

    @property
    def item_id(self) -> str | None:
        """Return the item ID (GUID) for this custom XML part.

        The GUID is used to identify this custom XML part when binding
        content controls to it. Format: "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}"
        """
        if self._properties_part is None:
            return None
        return self._properties_part.item_id

    @property
    def schema_uris(self) -> List[str]:
        """Return list of XML schema URIs referenced by this custom XML.

        These are the schema namespaces that the custom XML content conforms to.
        """
        if self._properties_part is None:
            return []
        return self._properties_part.schema_uris

    @property
    def element(self) -> etree._Element | None:
        """Return the root XML element of the custom XML content.

        This is an lxml Element that can be used for direct XML manipulation.
        """
        return self._xml_part.element

    @property
    def xml(self) -> bytes:
        """Return the raw XML content as bytes."""
        return self._xml_part.blob

    @property
    def partname(self) -> str:
        """Return the part name (path within the package)."""
        return str(self._xml_part.partname)

    def xpath(self, xpath_expr: str, namespaces: dict[str, str] | None = None) -> List[Any]:
        """Execute an XPath query on the custom XML content.

        Args:
            xpath_expr: XPath expression to evaluate.
            namespaces: Optional dict mapping prefixes to namespace URIs.

        Returns:
            List of matching elements or attribute values.

        Example::

            # Find all customer names
            names = custom_xml.xpath("//customer/name/text()")

            # With namespace
            items = custom_xml.xpath(
                "//ns:item",
                namespaces={"ns": "http://example.com/items"}
            )
        """
        return self._xml_part.xpath(xpath_expr, namespaces)

    def get_text(self, xpath_expr: str, namespaces: dict[str, str] | None = None) -> str | None:
        """Get the text content of the first element matching an XPath expression.

        Args:
            xpath_expr: XPath expression to evaluate.
            namespaces: Optional dict mapping prefixes to namespace URIs.

        Returns:
            Text content of the first matching element, or None if not found.

        Example::

            name = custom_xml.get_text("/customer/name")
        """
        results: List[Any] = self.xpath(xpath_expr, namespaces)
        if not results:
            return None
        result: Any = results[0]
        if isinstance(result, str):
            return result
        if hasattr(result, "text"):
            return result.text
        return str(result)

    def set_text(
        self, xpath_expr: str, value: str, namespaces: dict[str, str] | None = None
    ) -> bool:
        """Set the text content of elements matching an XPath expression.

        Args:
            xpath_expr: XPath expression to evaluate.
            value: New text value to set.
            namespaces: Optional dict mapping prefixes to namespace URIs.

        Returns:
            True if any elements were updated, False otherwise.

        Example::

            custom_xml.set_text("/customer/name", "ACME Corporation")
        """
        results: List[Any] = self.xpath(xpath_expr, namespaces)
        if not results:
            return False
        for result in results:
            if hasattr(result, "text"):
                result.text = value
        return True
