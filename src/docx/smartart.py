"""Proxy objects for SmartArt diagrams.

SmartArt diagrams are complex DrawingML objects that consist of multiple
interconnected parts (data, layout, colors, quickStyle).
"""

# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from lxml import etree

if TYPE_CHECKING:
    pass


class SmartArtCollection:
    """Collection of SmartArt diagrams in a document.

    Provides iteration and indexing over SmartArt diagrams found in the document.
    """

    def __init__(self, smartart_elements: List[etree._Element]):
        self._smartart_elements = smartart_elements

    def __iter__(self) -> Iterator[SmartArt]:
        """Iterate over all SmartArt diagrams."""
        for elem in self._smartart_elements:
            yield SmartArt(elem)

    def __len__(self) -> int:
        """Return the number of SmartArt diagrams."""
        return len(self._smartart_elements)

    def __getitem__(self, index: int) -> SmartArt:
        """Return SmartArt at `index`."""
        return SmartArt(self._smartart_elements[index])

    def __bool__(self) -> bool:
        """Return True if there are any SmartArt diagrams."""
        return len(self._smartart_elements) > 0


class SmartArt:
    """Proxy for a single SmartArt diagram.

    Provides read-only access to SmartArt text content extracted from the
    data model.
    """

    def __init__(self, element: etree._Element):
        self._element = element  # This is the dgm:relIds element or data model

    @property
    def text_content(self) -> List[str]:
        """Return list of text strings from the SmartArt nodes.

        SmartArt text is stored in dgm:pt (point) elements within the data model.
        This extracts all text from all nodes in the diagram.
        """
        texts: List[str] = []

        # SmartArt text is in a:t elements within dgm:t within dgm:pt
        dgm_ns = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
        a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        # Find all dgm:pt elements (data points/nodes)
        for pt in self._element.iter(f"{dgm_ns}pt"):
            # Find dgm:t (text container) within the point
            t_container = pt.find(f"{dgm_ns}t")
            if t_container is not None:
                # Find all a:t text elements
                for t_elem in t_container.iter(f"{a_ns}t"):
                    if t_elem.text:
                        texts.append(str(t_elem.text))

        return texts

    @property
    def text(self) -> str:
        """Return all SmartArt text as a single string, joined by newlines."""
        return "\n".join(self.text_content)

    @property
    def node_count(self) -> int:
        """Return the number of content nodes in the SmartArt.

        This counts dgm:pt elements that are not the document root (type='doc').
        """
        dgm_ns = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
        count = 0
        for pt in self._element.iter(f"{dgm_ns}pt"):
            pt_type = pt.get("type")
            if pt_type != "doc":  # Skip document root
                count += 1
        return count

    @property
    def nodes(self) -> List[SmartArtNode]:
        """Return list of content nodes in the SmartArt.

        Each node represents a shape or text container in the diagram.
        """
        dgm_ns = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
        nodes: List[SmartArtNode] = []
        for pt in self._element.iter(f"{dgm_ns}pt"):
            pt_type = pt.get("type")
            if pt_type != "doc":  # Skip document root
                nodes.append(SmartArtNode(pt))
        return nodes


class SmartArtNode:
    """Proxy for a single node in a SmartArt diagram.

    Provides access to and modification of node text content.
    """

    def __init__(self, element: etree._Element):
        self._element = element
        self._dgm_ns = "{http://schemas.openxmlformats.org/drawingml/2006/diagram}"
        self._a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    @property
    def model_id(self) -> str | None:
        """Return the model ID of this node.

        The model ID is used to identify the node within the diagram.
        """
        return self._element.get("modelId")

    @property
    def node_type(self) -> str | None:
        """Return the type of this node (e.g., 'node', 'sibTrans', 'parTrans').

        - 'node': A regular content node
        - 'sibTrans': Sibling transition (connector between siblings)
        - 'parTrans': Parent transition (connector to parent)
        """
        return self._element.get("type")

    @property
    def text(self) -> str:
        """Return the text content of this node.

        Returns all text from a:t elements within the node's dgm:t container.
        """
        t_container = self._element.find(f"{self._dgm_ns}t")
        if t_container is None:
            return ""

        texts: List[str] = []
        for t_elem in t_container.iter(f"{self._a_ns}t"):
            if t_elem.text:
                texts.append(str(t_elem.text))

        return "".join(texts)

    @text.setter
    def text(self, value: str) -> None:
        """Set the text content of this node.

        This replaces all existing text in the node with the new value.
        The text is placed in the first a:t element found, or creates one
        if none exists.

        Note: For complex text formatting, modify the element directly.
        """
        t_container = self._element.find(f"{self._dgm_ns}t")

        if t_container is None:
            # Create the text container structure
            t_container = etree.SubElement(
                self._element,
                f"{self._dgm_ns}t"
            )

        # Find or create the body element (a:bodyPr is optional but a:p is needed)
        p_elem = t_container.find(f".//{self._a_ns}p")
        if p_elem is None:
            # Create minimal paragraph structure
            p_elem = etree.SubElement(t_container, f"{self._a_ns}p")

        # Find or create the run
        r_elem = p_elem.find(f"{self._a_ns}r")
        if r_elem is None:
            r_elem = etree.SubElement(p_elem, f"{self._a_ns}r")

        # Find or create the text element
        t_elem = r_elem.find(f"{self._a_ns}t")
        if t_elem is None:
            t_elem = etree.SubElement(r_elem, f"{self._a_ns}t")

        # Set the text
        t_elem.text = value

        # Clear any other text elements to avoid duplication
        for other_t in t_container.iter(f"{self._a_ns}t"):
            if other_t is not t_elem:
                other_t.text = ""
