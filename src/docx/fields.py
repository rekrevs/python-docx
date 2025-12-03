"""Fields proxy objects for simple and complex fields."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from docx.oxml.fields import ComplexField, CT_FldSimple


class Fields:
    """Collection of fields in a document.

    This collection provides access to both simple fields (`<w:fldSimple>`)
    and complex fields (delimited by `<w:fldChar>` markers).

    Fields are used for dynamic content like:
    - Page numbers (PAGE, NUMPAGES)
    - Dates and times (DATE, TIME)
    - Document properties (TITLE, AUTHOR)
    - Cross-references (REF, PAGEREF)
    - Table of contents (TOC)
    - Citations and bibliographies (CITATION)
    - Mail merge fields (MERGEFIELD)
    """

    def __init__(
        self,
        simple_fields: list[CT_FldSimple],
        complex_fields: list[ComplexField],
    ) -> None:
        self._simple_fields = simple_fields
        self._complex_fields = complex_fields

    def __iter__(self) -> Iterator[SimpleField | ComplexFieldProxy]:
        """Iterate over all fields (simple and complex) in the document."""
        for sf in self._simple_fields:
            yield SimpleField(sf)
        for cf in self._complex_fields:
            yield ComplexFieldProxy(cf)

    def __len__(self) -> int:
        """Return the total number of fields."""
        return len(self._simple_fields) + len(self._complex_fields)

    @property
    def simple(self) -> list[SimpleField]:
        """List of simple fields in the document."""
        return [SimpleField(sf) for sf in self._simple_fields]

    @property
    def complex(self) -> list[ComplexFieldProxy]:
        """List of complex fields in the document."""
        return [ComplexFieldProxy(cf) for cf in self._complex_fields]

    def filter_by_type(self, field_type: str) -> Iterator[SimpleField | ComplexFieldProxy]:
        """Yield fields of the specified type.

        `field_type` is case-insensitive and should be the field name like
        'PAGE', 'DATE', 'TOC', 'REF', 'CITATION', etc.
        """
        field_type_upper = field_type.upper()
        for field in self:
            if field.field_type.upper() == field_type_upper:
                yield field


class SimpleField:
    """Proxy for a simple field (`<w:fldSimple>` element).

    Simple fields contain the field code in an attribute and display
    their result as child content. They are used for basic fields like
    page numbers that don't require complex formatting.

    Example XML:
        <w:fldSimple w:instr=" PAGE ">
            <w:r><w:t>1</w:t></w:r>
        </w:fldSimple>
    """

    def __init__(self, fld_simple: CT_FldSimple) -> None:
        self._fld_simple = fld_simple

    @property
    def field_code(self) -> str:
        """The complete field instruction code.

        This includes the field type and any switches/parameters.
        Example: 'PAGE \\* MERGEFORMAT' or 'DATE \\@ "MMMM d, yyyy"'
        """
        return self._fld_simple.field_code

    @property
    def field_type(self) -> str:
        """The field type (first word of the instruction).

        Common types include:
        - PAGE: Current page number
        - NUMPAGES: Total page count
        - DATE: Current date
        - TIME: Current time
        - TITLE: Document title
        - AUTHOR: Document author
        """
        return self._fld_simple.field_type

    @property
    def result(self) -> str:
        """The cached result text of this field.

        This is the text currently displayed in the document. Note that
        this may be stale if the field hasn't been updated since the
        document was last opened in Word.
        """
        return self._fld_simple.result_text

    @property
    def is_simple(self) -> bool:
        """Always True for simple fields."""
        return True

    @property
    def is_complex(self) -> bool:
        """Always False for simple fields."""
        return False


class ComplexFieldProxy:
    """Proxy for a complex field.

    Complex fields are represented by a series of elements:
    - `<w:fldChar w:fldCharType="begin"/>` marks the start
    - `<w:instrText>` elements contain the field code
    - `<w:fldChar w:fldCharType="separate"/>` separates code from result
    - Result content (runs with text)
    - `<w:fldChar w:fldCharType="end"/>` marks the end

    Complex fields are used when:
    - The field code spans multiple runs
    - Different parts of the result need different formatting
    - The field is more complex (TOC, cross-references, citations)

    Example field codes:
        - TOC \\o "1-3" \\h \\z \\u (Table of contents)
        - PAGEREF _Toc123 \\h (Page reference to bookmark)
        - CITATION Smith2020 \\l 1033 (Citation)
        - REF _Ref456 \\h (Cross-reference to bookmark)
    """

    def __init__(self, complex_field: ComplexField) -> None:
        self._complex_field = complex_field

    @property
    def field_code(self) -> str:
        """The complete field instruction code.

        This includes the field type and any switches/parameters.
        Multiple instrText elements are combined into a single string.
        """
        return self._complex_field.field_code

    @property
    def field_type(self) -> str:
        """The field type (first word of the instruction).

        Common complex field types include:
        - TOC: Table of contents
        - REF: Cross-reference
        - PAGEREF: Page reference
        - CITATION: Bibliography citation
        - HYPERLINK: Hyperlink
        - SEQ: Sequence number
        - IF: Conditional field
        """
        return self._complex_field.field_type

    @property
    def result(self) -> str:
        """The cached result text of this field.

        This is the text currently displayed in the document between
        the separator and end markers. Note that this may be stale if
        the field hasn't been updated.
        """
        return self._complex_field.result_text

    @property
    def is_simple(self) -> bool:
        """Always False for complex fields."""
        return False

    @property
    def is_complex(self) -> bool:
        """Always True for complex fields."""
        return True

    @property
    def is_complete(self) -> bool:
        """True if this field has matching begin and end markers.

        Incomplete fields (missing end marker) may occur in malformed
        documents and their behavior is undefined.
        """
        return self._complex_field.is_complete
