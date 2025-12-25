"""Custom element classes for Custom XML parts (ds: namespace)."""

# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false

from __future__ import annotations

from typing import List

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne


class CT_DatastoreItem(BaseOxmlElement):
    """``<ds:datastoreItem>`` element, the root element of a custom XML properties part.

    Contains the itemID (GUID) and schema references for a custom XML part.
    """

    _schemaRefs: CT_DatastoreSchemaRefs | None = ZeroOrOne("ds:schemaRefs")  # type: ignore[assignment]

    @property
    def item_id(self) -> str | None:
        """Return the itemID attribute value (GUID) or None if not present."""
        return self.get(qn("ds:itemID"))

    @item_id.setter
    def item_id(self, value: str) -> None:
        """Set the itemID attribute value."""
        self.set(qn("ds:itemID"), value)

    @property
    def schema_uris(self) -> List[str]:
        """Return a list of schema URIs from all schemaRef elements."""
        uris = []
        schema_refs = self._schemaRefs
        if schema_refs is not None:
            for schema_ref in schema_refs.schemaRef_lst:
                uri = schema_ref.uri
                if uri:
                    uris.append(uri)
        return uris

    @classmethod
    def new(cls, item_id: str) -> CT_DatastoreItem:
        """Create a new ds:datastoreItem element with the given itemID.

        Args:
            item_id: The GUID for the custom XML part (including braces).

        Returns:
            A new CT_DatastoreItem element.
        """
        return parse_xml(  # type: ignore[return-value]
            f'<ds:datastoreItem {nsdecls("ds")} ds:itemID="{item_id}"/>'
        )


class CT_DatastoreSchemaRefs(BaseOxmlElement):
    """``<ds:schemaRefs>`` element, contains schema references for a custom XML part."""

    # Type annotation for metaclass-generated property
    schemaRef_lst: List[CT_DatastoreSchemaRef]

    schemaRef = ZeroOrMore("ds:schemaRef")


class CT_DatastoreSchemaRef(BaseOxmlElement):
    """``<ds:schemaRef>`` element, a single schema reference."""

    @property
    def uri(self) -> str | None:
        """Return the schema URI."""
        return self.get(qn("ds:uri"))

    @uri.setter
    def uri(self, value: str) -> None:
        """Set the schema URI."""
        self.set(qn("ds:uri"), value)
