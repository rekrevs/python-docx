"""Proxy objects for Office Math Markup Language (OMML) equations.

Math equations in Word documents use OMML, a specialized XML vocabulary
for representing mathematical expressions.
"""

# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from lxml import etree

if TYPE_CHECKING:
    from docx.oxml.math import CT_OMath, CT_OMathPara


class MathEquations:
    """Collection of math equations in a document.

    Provides iteration and indexing over all math zones (m:oMath and m:oMathPara)
    found in the document.
    """

    def __init__(self, omath_elements: List[CT_OMath], omathpara_elements: List[CT_OMathPara]):
        self._omath_elements = omath_elements
        self._omathpara_elements = omathpara_elements

    def __iter__(self) -> Iterator[MathEquation]:
        """Iterate over all math equations."""
        # First yield inline math (m:oMath not in m:oMathPara)
        for elem in self._omath_elements:
            yield MathEquation(elem, is_block=False)
        # Then yield block math (m:oMathPara)
        for elem in self._omathpara_elements:
            yield MathEquation(elem, is_block=True)

    def __len__(self) -> int:
        """Return the total number of math equations."""
        return len(self._omath_elements) + len(self._omathpara_elements)

    def __getitem__(self, index: int) -> MathEquation:
        """Return MathEquation at `index`."""
        total_inline = len(self._omath_elements)
        if index < total_inline:
            return MathEquation(self._omath_elements[index], is_block=False)
        return MathEquation(self._omathpara_elements[index - total_inline], is_block=True)

    @property
    def inline(self) -> List[MathEquation]:
        """Return list of inline math equations (m:oMath not in m:oMathPara)."""
        return [MathEquation(elem, is_block=False) for elem in self._omath_elements]

    @property
    def block(self) -> List[MathEquation]:
        """Return list of block-level math equations (m:oMathPara)."""
        return [MathEquation(elem, is_block=True) for elem in self._omathpara_elements]


class MathEquation:
    """Proxy for a single math equation.

    Provides access to the math content and metadata.
    """

    def __init__(self, element: CT_OMath | CT_OMathPara, is_block: bool):
        self._element = element
        self._is_block = is_block

    @property
    def is_block(self) -> bool:
        """Return True if this is a block-level equation (m:oMathPara).

        Block-level equations appear on their own line, often centered.
        Inline equations appear within the text flow.
        """
        return self._is_block

    @property
    def is_inline(self) -> bool:
        """Return True if this is an inline equation."""
        return not self._is_block

    @property
    def text(self) -> str:
        """Return a plain text approximation of the math content.

        This provides a human-readable representation of the equation,
        though it may lose some mathematical structure. Useful for
        searching or basic text extraction.

        Examples of text representation:
        - Fraction a/b: "(a/b)"
        - Superscript x^2: "x^2"
        - Subscript x_i: "x_i"
        - Square root: "sqrt(x)"
        - Summation: "∑[i=1,n](x_i)"
        """
        return self._element.math_text

    @property
    def xml(self) -> bytes:
        """Return the raw OMML XML for this equation."""
        return etree.tostring(self._element, encoding="UTF-8")

    @property
    def element(self) -> etree._Element:
        """Return the underlying lxml element for direct manipulation."""
        return self._element
