"""Custom element classes related to fields (simple and complex)."""

from __future__ import annotations

from enum import Enum
from typing import TYPE_CHECKING, List, cast

from docx.oxml.ns import qn
from docx.oxml.simpletypes import ST_String
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute

if TYPE_CHECKING:
    from lxml.etree import _Element  # pyright: ignore[reportPrivateUsage]

    from docx.oxml.document import CT_Body


class ST_FldCharType(Enum):
    """Enumeration for fldCharType attribute values."""
    BEGIN = "begin"
    SEPARATE = "separate"
    END = "end"


class CT_FldSimple(BaseOxmlElement):
    """`w:fldSimple` element for simple fields.

    A simple field contains the field code in its `instr` attribute and
    displays the field result in its content (typically a run).

    Example:
        <w:fldSimple w:instr="PAGE">
            <w:r><w:t>1</w:t></w:r>
        </w:fldSimple>
    """

    instr: str = RequiredAttribute("w:instr", ST_String)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(cls, field_type: str, switches: str = "", result: str = "") -> CT_FldSimple:
        """Create a new `w:fldSimple` element.

        Args:
            field_type: The field type (e.g., 'PAGE', 'DATE', 'NUMPAGES').
            switches: Optional field switches (e.g., r'\\@ "MMMM d, yyyy"').
            result: Optional placeholder result text (displayed until updated).

        Returns:
            A new CT_FldSimple element with a child run containing the result.
        """
        from docx.oxml.parser import OxmlElement

        # Build instruction string
        instr = f" {field_type}"
        if switches:
            instr = f"{instr} {switches}"
        instr = f"{instr} "

        # Create fldSimple element
        fld_simple = OxmlElement("w:fldSimple")
        fld_simple.set(qn("w:instr"), instr)

        # Create child run with result text
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = result if result else ""
        r.append(t)
        fld_simple.append(r)

        return cast("CT_FldSimple", fld_simple)

    @property
    def field_code(self) -> str:
        """The field instruction code (e.g., 'PAGE', 'DATE', 'NUMPAGES')."""
        return self.instr.strip()

    @property
    def field_type(self) -> str:
        """The field type (first word of the instruction).

        For example, 'PAGE' for a page number field, 'DATE' for a date field.
        """
        return self.field_code.split()[0] if self.field_code else ""

    @property
    def result_text(self) -> str:
        """The cached result text of this field.

        This is the text displayed in the document. It may be stale if the
        field hasn't been updated.
        """
        # Get all text from descendant w:t elements
        text_elements = self.xpath(".//w:t/text()")
        return "".join(text_elements)


class CT_FldChar(BaseOxmlElement):
    """`w:fldChar` element marking parts of a complex field.

    Complex fields consist of:
    - A begin marker (fldCharType="begin")
    - Field instructions in w:instrText elements
    - A separator marker (fldCharType="separate")
    - Field result (displayed text)
    - An end marker (fldCharType="end")
    """

    fldCharType: str = RequiredAttribute("w:fldCharType", ST_String)  # pyright: ignore[reportAssignmentType]

    @property
    def char_type(self) -> ST_FldCharType:
        """The type of field character (begin, separate, or end)."""
        return ST_FldCharType(self.fldCharType)

    @property
    def is_begin(self) -> bool:
        """True if this is a field begin marker."""
        return self.fldCharType == "begin"

    @property
    def is_separate(self) -> bool:
        """True if this is a field separator marker."""
        return self.fldCharType == "separate"

    @property
    def is_end(self) -> bool:
        """True if this is a field end marker."""
        return self.fldCharType == "end"


class CT_FldInstrText(BaseOxmlElement):
    """`w:instrText` element containing field instruction text.

    This element contains part of or all of the field instruction code
    for a complex field. Multiple instrText elements may be combined
    to form the complete field instruction.
    """

    @property
    def instr_text(self) -> str:
        """The instruction text content."""
        # Use lxml's text property directly - BaseOxmlElement inherits from _Element
        text = super().text
        return text if text is not None else ""


class ComplexField:
    """Represents a complex field parsed from document XML.

    A complex field is delimited by fldChar elements and contains:
    - Field instruction code (from instrText elements)
    - Field result (the displayed text between separator and end)

    This is a data class for representing parsed field information,
    not an OXML element class.
    """

    def __init__(
        self,
        begin_elem: CT_FldChar,
        instruction: str,
        result: str,
        end_elem: CT_FldChar | None = None,
    ):
        self._begin_elem = begin_elem
        self._instruction = instruction
        self._result = result
        self._end_elem = end_elem

    @property
    def field_code(self) -> str:
        """The complete field instruction code."""
        return self._instruction.strip()

    @property
    def field_type(self) -> str:
        """The field type (first word of instruction).

        Examples: 'PAGE', 'DATE', 'TOC', 'REF', 'CITATION', 'PAGEREF'
        """
        code = self.field_code
        return code.split()[0] if code else ""

    @property
    def result_text(self) -> str:
        """The cached result text of this field."""
        return self._result

    @property
    def is_complete(self) -> bool:
        """True if this field has both begin and end markers."""
        return self._end_elem is not None


def iter_complex_fields(body_element: CT_Body) -> List[ComplexField]:
    """Parse and yield all complex fields in a document body element.

    Complex fields are identified by matching begin/end fldChar markers
    and collecting the instrText elements between them.

    Args:
        body_element: The CT_Body element to search for fields.

    Returns:
        List of ComplexField objects representing each field found.
    """
    fields: List[ComplexField] = []
    field_stack: List[tuple[_Element, str, str, bool, bool]] = []  # Stack for nested fields

    # Get all runs in document order
    all_runs: List[_Element] = body_element.xpath(".//w:r")

    current_instruction: List[str] = []
    current_result: List[str] = []
    in_instruction = False
    in_result = False
    current_begin: _Element | None = None

    for run in all_runs:
        # Check for fldChar
        fld_chars: List[_Element] = run.xpath("./w:fldChar")
        for fld_char in fld_chars:
            char_type = fld_char.get(
                "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fldCharType"
            )

            if char_type == "begin":
                # Start of a new field (possibly nested)
                if current_begin is not None:
                    # Push current field onto stack (nested field)
                    field_stack.append((
                        current_begin,
                        "".join(current_instruction),
                        "".join(current_result),
                        in_instruction,
                        in_result,
                    ))
                current_begin = fld_char
                current_instruction = []
                current_result = []
                in_instruction = True
                in_result = False

            elif char_type == "separate":
                in_instruction = False
                in_result = True

            elif char_type == "end":
                if current_begin is not None:
                    # Create field object
                    field = ComplexField(
                        begin_elem=cast(CT_FldChar, current_begin),
                        instruction="".join(current_instruction),
                        result="".join(current_result),
                        end_elem=cast(CT_FldChar, fld_char),
                    )
                    fields.append(field)

                    # Pop parent field from stack if any
                    if field_stack:
                        (
                            current_begin,
                            prev_instr,
                            prev_result,
                            in_instruction,
                            in_result,
                        ) = field_stack.pop()
                        current_instruction = [prev_instr]
                        current_result = [prev_result]
                    else:
                        current_begin = None
                        current_instruction = []
                        current_result = []
                        in_instruction = False
                        in_result = False

        # Collect instruction text
        if in_instruction:
            instr_texts: List[str] = run.xpath("./w:instrText/text()")
            current_instruction.extend(instr_texts)

        # Collect result text
        if in_result:
            result_texts: List[str] = run.xpath("./w:t/text()")
            current_result.extend(result_texts)

    return fields
