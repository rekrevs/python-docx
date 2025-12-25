# pyright: reportImportCycles=false
# pyright: reportPrivateUsage=false

"""|Document| and closely related objects."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, Iterator, List, Sequence

from docx.blkcntnr import BlockItemContainer
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_BREAK
from docx.section import Section, Sections
from docx.shared import ElementProxy, Emu, Inches, Length
from docx.text.run import Run

if TYPE_CHECKING:
    import docx.types as t
    from docx.bookmarks import Bookmark, Bookmarks
    from docx.chart import Charts
    from docx.comments import Comment, Comments
    from docx.customxml import CustomXml, CustomXmlParts
    from docx.fields import Fields
    from docx.footnotes import Endnotes, Footnotes
    from docx.math import MathEquations
    from docx.oxml.document import CT_Body, CT_Document
    from docx.parts.document import DocumentPart
    from docx.revisions import Revisions
    from docx.sdt import SdtBlockContentControl, SdtContentControls
    from docx.settings import Settings
    from docx.smartart import SmartArtCollection
    from docx.styles.style import ParagraphStyle, _TableStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.textbox import TextBox, TextBoxes


class Document(ElementProxy):
    """WordprocessingML (WML) document.

    Not intended to be constructed directly. Use :func:`docx.Document` to open or create
    a document.
    """

    def __init__(self, element: CT_Document, part: DocumentPart):
        super(Document, self).__init__(element)
        self._element = element
        self._part = part
        self.__body = None

    def add_comment(
        self,
        runs: Run | Sequence[Run],
        text: str | None = "",
        author: str = "",
        initials: str | None = "",
    ) -> Comment:
        """Add a comment to the document, anchored to the specified runs.

        `runs` can be a single `Run` object or a non-empty sequence of `Run` objects. Only the
        first and last run of a sequence are used, it's just more convenient to pass a whole
        sequence when that's what you have handy, like `paragraph.runs` for example. When `runs`
        contains a single `Run` object, that run serves as both the first and last run.

        A comment can be anchored only on an even run boundary, meaning the text the comment
        "references" must be a non-zero integer number of consecutive runs. The runs need not be
        _contiguous_ per se, like the first can be in one paragraph and the last in the next
        paragraph, but all runs between the first and the last will be included in the reference.

        The comment reference range is delimited by placing a `w:commentRangeStart` element before
        the first run and a `w:commentRangeEnd` element after the last run. This is why only the
        first and last run are required and why a single run can serve as both first and last.
        Word works out which text to highlight in the UI based on these range markers.

        `text` allows the contents of a simple comment to be provided in the call, providing for
        the common case where a comment is a single phrase or sentence without special formatting
        such as bold or italics. More complex comments can be added using the returned `Comment`
        object in much the same way as a `Document` or (table) `Cell` object, using methods like
        `.add_paragraph()`, .add_run()`, etc.

        The `author` and `initials` parameters allow that metadata to be set for the comment.
        `author` is a required attribute on a comment and is the empty string by default.
        `initials` is optional on a comment and may be omitted by passing |None|, but Word adds an
        `initials` attribute by default and we follow that convention by using the empty string
        when no `initials` argument is provided.
        """
        # -- normalize `runs` to a sequence of runs --
        runs = [runs] if isinstance(runs, Run) else runs
        first_run = runs[0]
        last_run = runs[-1]

        # -- Note that comments can only appear in the document part --
        comment = self.comments.add_comment(text=text, author=author, initials=initials)

        # -- let the first run orchestrate placement of the comment range start and end --
        first_run.mark_comment_range(last_run, comment.comment_id)

        return comment

    def add_heading(self, text: str = "", level: int = 1):
        """Return a heading paragraph newly added to the end of the document.

        The heading paragraph will contain `text` and have its paragraph style
        determined by `level`. If `level` is 0, the style is set to `Title`. If `level`
        is 1 (or omitted), `Heading 1` is used. Otherwise the style is set to `Heading
        {level}`. Raises |ValueError| if `level` is outside the range 0-9.
        """
        if not 0 <= level <= 9:
            raise ValueError("level must be in range 0-9, got %d" % level)
        style = "Title" if level == 0 else "Heading %d" % level
        return self.add_paragraph(text, style)

    def add_page_break(self):
        """Return newly |Paragraph| object containing only a page break."""
        paragraph = self.add_paragraph()
        paragraph.add_run().add_break(WD_BREAK.PAGE)
        return paragraph

    def add_paragraph(self, text: str = "", style: str | ParagraphStyle | None = None) -> Paragraph:
        """Return paragraph newly added to the end of the document.

        The paragraph is populated with `text` and having paragraph style `style`.

        `text` can contain tab (``\\t``) characters, which are converted to the
        appropriate XML form for a tab. `text` can also include newline (``\\n``) or
        carriage return (``\\r``) characters, each of which is converted to a line
        break.
        """
        return self._body.add_paragraph(text, style)

    def add_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
    ):
        """Return new picture shape added in its own paragraph at end of the document.

        The picture contains the image at `image_path_or_stream`, scaled based on
        `width` and `height`. If neither width nor height is specified, the picture
        appears at its native size. If only one is specified, it is used to compute a
        scaling factor that is then applied to the unspecified dimension, preserving the
        aspect ratio of the image. The native size of the picture is calculated using
        the dots-per-inch (dpi) value specified in the image file, defaulting to 72 dpi
        if no value is specified, as is often the case.
        """
        run = self.add_paragraph().add_run()
        return run.add_picture(image_path_or_stream, width, height)

    def add_floating_picture(
        self,
        image_path_or_stream: str | IO[bytes],
        width: int | Length | None = None,
        height: int | Length | None = None,
        pos_x: int | Length = Emu(0),
        pos_y: int | Length = Emu(0),
        behind_doc: bool = False,
        wrap_type: str = "square",
        h_relative_from: str = "column",
        v_relative_from: str = "paragraph",
    ):
        """Return new floating picture shape added in its own paragraph.

        The picture is positioned as an anchored/floating shape rather than inline with text.

        Args:
            image_path_or_stream: Path or file-like object containing the image.
            width: Width of the image (None for native size).
            height: Height of the image (None for native size).
            pos_x: Horizontal position offset in EMUs (or Inches, Pt, etc.).
            pos_y: Vertical position offset in EMUs (or Inches, Pt, etc.).
            behind_doc: If True, image is placed behind document text.
            wrap_type: Text wrapping style. One of:
                - 'none': No wrapping (image floats over text)
                - 'square': Square wrapping
                - 'tight': Tight wrapping
                - 'through': Through wrapping
                - 'topAndBottom': Text flows above and below only
            h_relative_from: Horizontal position relative to. One of:
                - 'character', 'column', 'insideMargin', 'leftMargin', 'margin',
                  'outsideMargin', 'page', 'rightMargin'
            v_relative_from: Vertical position relative to. One of:
                - 'insideMargin', 'line', 'margin', 'outsideMargin', 'page',
                  'paragraph', 'topMargin', 'bottomMargin'

        Returns:
            FloatingShape: The newly created floating picture shape.

        Example::

            from docx.shared import Inches

            # Add floating image at specific position
            shape = document.add_floating_picture(
                'image.png',
                width=Inches(2),
                pos_x=Inches(1),
                pos_y=Inches(2),
                wrap_type='square'
            )

            # Add image behind text (e.g., watermark)
            shape = document.add_floating_picture(
                'watermark.png',
                behind_doc=True,
                wrap_type='none'
            )
        """
        from docx.shape import FloatingShape

        # Convert pos_x and pos_y to int if they are Length objects
        px = int(pos_x) if hasattr(pos_x, '__int__') else pos_x
        py = int(pos_y) if hasattr(pos_y, '__int__') else pos_y

        anchor = self._part.new_pic_anchor(
            image_path_or_stream, width, height,
            Emu(px), Emu(py),
            behind_doc=behind_doc,
            wrap_type=wrap_type,
            h_relative_from=h_relative_from,
            v_relative_from=v_relative_from,
        )
        run = self.add_paragraph().add_run()
        run._r.add_drawing(anchor)
        return FloatingShape(anchor, self._part)

    def add_text_box(
        self,
        width: int | Length,
        height: int | Length,
        pos_x: int | Length = Emu(0),
        pos_y: int | Length = Emu(0),
        wrap_type: str = "square",
    ) -> "TextBox":
        """Return a new text box added in its own paragraph.

        The text box is positioned as a floating shape and can contain
        paragraphs and tables.

        Args:
            width: Width of the text box in EMUs (or Inches, Pt, etc.).
            height: Height of the text box in EMUs (or Inches, Pt, etc.).
            pos_x: Horizontal position offset in EMUs (default 0).
            pos_y: Vertical position offset in EMUs (default 0).
            wrap_type: Text wrapping style. One of:
                - 'none': No wrapping (text box floats over text)
                - 'square': Square wrapping
                - 'topAndBottom': Text flows above and below only

        Returns:
            TextBox: The newly created text box.

        Example::

            from docx.shared import Inches

            # Add a text box
            text_box = document.add_text_box(
                width=Inches(3),
                height=Inches(2),
                pos_x=Inches(1),
                pos_y=Inches(1),
            )

            # Add content to the text box
            text_box.add_paragraph('Hello from the text box!')
            text_box.add_paragraph('This is a second paragraph.')
        """
        from docx.oxml.mce import CT_AlternateContent
        from docx.textbox import TextBox

        # Get next shape ID
        shape_id = self._part.next_id

        # Convert to int if Length objects
        cx = int(width) if hasattr(width, "__int__") else width
        cy = int(height) if hasattr(height, "__int__") else height
        px = int(pos_x) if hasattr(pos_x, "__int__") else pos_x
        py = int(pos_y) if hasattr(pos_y, "__int__") else pos_y

        # Create the text box XML structure
        ac = CT_AlternateContent.new_textbox(
            shape_id=shape_id,
            cx=Emu(cx),
            cy=Emu(cy),
            pos_x=Emu(px),
            pos_y=Emu(py),
            wrap_type=wrap_type,
        )

        # Add to a paragraph run
        run = self.add_paragraph().add_run()
        run._r.append(ac)

        return TextBox(ac, self._part)

    def add_section(self, start_type: WD_SECTION = WD_SECTION.NEW_PAGE):
        """Return a |Section| object newly added at the end of the document.

        The optional `start_type` argument must be a member of the :ref:`WdSectionStart`
        enumeration, and defaults to ``WD_SECTION.NEW_PAGE`` if not provided.
        """
        new_sectPr = self._element.body.add_section_break()
        new_sectPr.start_type = start_type
        return Section(new_sectPr, self._part)

    def add_table(self, rows: int, cols: int, style: str | _TableStyle | None = None):
        """Add a table having row and column counts of `rows` and `cols` respectively.

        `style` may be a table style object or a table style name. If `style` is |None|,
        the table inherits the default table style of the document.
        """
        table = self._body.add_table(rows, cols, self._block_width)
        table.style = style
        return table

    @property
    def bookmarks(self) -> Bookmarks:
        """A |Bookmarks| collection providing access to bookmarks in this document.

        Bookmarks are named ranges that can be used as targets for cross-references,
        hyperlinks, or programmatic document navigation.

        Example::

            # Iterate over all bookmarks
            for bookmark in document.bookmarks:
                print(f"{bookmark.name}: id={bookmark.bookmark_id}")

            # Get bookmark by name
            toc_bm = document.bookmarks.get("_Toc123456789")

            # Check if bookmark exists
            if "MyBookmark" in document.bookmarks:
                print("Found it!")

            # Filter system vs user bookmarks
            user_bookmarks = [bm for bm in document.bookmarks if not bm.is_system]
        """
        from docx.bookmarks import Bookmarks

        bookmark_starts = self._element.body.xpath(".//w:bookmarkStart")
        return Bookmarks(bookmark_starts)

    def add_bookmark(
        self,
        name: str,
        start: Paragraph | None = None,
        end: Paragraph | None = None,
    ) -> Bookmark:
        """Add a bookmark to the document.

        Args:
            name: The name for the bookmark. Must be unique within the document.
            start: The paragraph to start the bookmark at. If None, creates a point
                   bookmark at the end of the document body.
            end: The paragraph to end the bookmark at. If None, uses the same
                 paragraph as start (bookmark spans just that paragraph).

        Returns:
            The newly created Bookmark.

        Raises:
            ValueError: If a bookmark with this name already exists.

        Example::

            # Bookmark a single paragraph
            para = document.paragraphs[0]
            bookmark = document.add_bookmark("chapter1", start=para)

            # Bookmark a range of paragraphs
            start_para = document.paragraphs[0]
            end_para = document.paragraphs[2]
            bookmark = document.add_bookmark("section1", start=start_para, end=end_para)

            # Create a point bookmark at end of document
            bookmark = document.add_bookmark("insert_point")
        """
        from docx.bookmarks import Bookmark
        from docx.oxml.bookmarks import CT_Bookmark, CT_MarkupRange

        # Validate name uniqueness
        if name in self.bookmarks:
            raise ValueError(f"A bookmark named '{name}' already exists")

        # Find next available bookmark ID
        bookmark_id = self._next_bookmark_id()

        # Create bookmark elements
        bookmark_start = CT_Bookmark.new(bookmark_id, name)
        bookmark_end = CT_MarkupRange.new(bookmark_id)

        # Place the bookmark elements
        body = self._element.body

        if start is None:
            # Point bookmark at end of body
            body.append(bookmark_start)
            body.append(bookmark_end)
        else:
            # Bookmark around paragraph(s)
            start_p = start._element
            end_p = end._element if end is not None else start_p

            # Insert bookmarkStart before start paragraph
            start_p.addprevious(bookmark_start)

            # Insert bookmarkEnd after end paragraph
            end_p.addnext(bookmark_end)

        return Bookmark(bookmark_start)

    def _next_bookmark_id(self) -> int:
        """Get the next available bookmark ID.

        Scans all existing bookmarkStart elements to find the maximum ID,
        then returns max + 1.
        """
        bookmark_starts = self._element.body.xpath(".//w:bookmarkStart")
        if not bookmark_starts:
            return 0

        max_id = max(
            int(bm.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id", "0"))
            for bm in bookmark_starts
        )
        return max_id + 1

    @property
    def charts(self) -> Charts:
        """A |Charts| collection of charts embedded in this document.

        Charts are DrawingML objects that display data visually. This collection
        provides read-only access to detect charts and examine their basic properties.

        Example::

            # Check for charts
            if document.charts:
                print(f"Found {len(document.charts)} charts")

            # Iterate over charts
            for chart in document.charts:
                print(f"Type: {chart.chart_type}")
                print(f"Title: {chart.title}")
                print(f"Series: {chart.series_count}")

        Note:
            This provides detection and basic metadata only. Full chart manipulation
            requires access to the embedded chart part and its data.
        """
        from docx.chart import Charts
        from docx.opc.constants import RELATIONSHIP_TYPE as RT

        body = self._element.body
        chart_elements = []

        # Find all c:chart references in drawings
        # Charts are referenced via r:id in a:graphicData with chart URI
        chart_refs = body.xpath(
            ".//c:chart",
        )

        # For each chart reference, try to get the actual chart part
        for chart_ref in chart_refs:
            # Get the r:id attribute
            r_id = chart_ref.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if r_id:
                try:
                    # Get the related chart part
                    chart_part = self._part.related_parts.get(r_id)
                    if chart_part is not None:
                        chart_elements.append(chart_part.element)
                except (KeyError, AttributeError):
                    pass

        return Charts(chart_elements)

    @property
    def comments(self) -> Comments:
        """A |Comments| object providing access to comments added to the document."""
        return self._part.comments

    @property
    def content_controls(self) -> SdtContentControls:
        """A |SdtContentControls| collection of content controls in this document.

        Content controls (also called Structured Document Tags or SDTs) are used for
        creating fillable forms, data-bound regions, and document building blocks.

        This property provides access to block-level content controls in the document
        body. Content controls in headers, footers, or table cells are not included.

        Content controls can be accessed by index, iterated, or searched by tag or alias.

        Example::

            # Iterate over all content controls
            for cc in document.content_controls:
                print(f"{cc.tag}: {cc.text}")

            # Find a specific content control by tag
            name_field = document.content_controls.get_by_tag("customer_name")
            if name_field:
                print(name_field.text)
        """
        from docx.sdt import SdtContentControls

        return SdtContentControls(self._element.body.sdt_lst, self._part)

    def add_content_control(
        self,
        sdt_type: str = "richText",
        tag: str | None = None,
        alias: str | None = None,
        placeholder_text: str = "",
    ) -> SdtBlockContentControl:
        """Add a content control (structured document tag) to the document body.

        Args:
            sdt_type: The type of content control. One of 'richText', 'text', 'date',
                     'dropDownList', 'comboBox'. Default is 'richText'.
            tag: Optional tag value for identifying the control programmatically.
            alias: Optional alias (title) displayed in the UI.
            placeholder_text: Optional placeholder text for the content.

        Returns:
            The newly created SdtBlockContentControl.

        Example::

            # Create a rich text content control
            cc = document.add_content_control(
                sdt_type="richText",
                tag="notes",
                alias="Notes",
                placeholder_text="Enter notes here"
            )

            # Create a dropdown list
            cc = document.add_content_control(
                sdt_type="dropDownList",
                tag="status",
                alias="Status"
            )
            cc.add_list_item("Draft", "draft")
            cc.add_list_item("Final", "final")

            # Create a date picker
            cc = document.add_content_control(
                sdt_type="date",
                tag="due_date",
                alias="Due Date"
            )
        """
        from docx.oxml.sdt import CT_SdtBlock
        from docx.sdt import SdtBlockContentControl

        sdt = CT_SdtBlock.new(sdt_type, tag, alias, placeholder_text)
        self._element.body.append(sdt)
        return SdtBlockContentControl(sdt, self._part)

    @property
    def core_properties(self):
        """A |CoreProperties| object providing Dublin Core properties of document."""
        return self._part.core_properties

    @property
    def custom_xml_parts(self) -> CustomXmlParts:
        """A |CustomXmlParts| collection of custom XML parts in this document.

        Custom XML parts store arbitrary XML data that can be bound to content
        controls for data-driven document generation.

        Example::

            # Iterate over custom XML parts
            for custom_xml in document.custom_xml_parts:
                print(f"Item ID: {custom_xml.item_id}")
                print(f"Schemas: {custom_xml.schema_uris}")

            # Get by item ID (GUID)
            cxml = document.custom_xml_parts.get_by_item_id("{C66AC4F1-97D3-...}")
            if cxml:
                name = cxml.get_text("/customer/name")

            # Get by root element namespace
            cxml = document.custom_xml_parts.get_by_namespace("http://example.com/schema")

            # Query with XPath
            for cxml in document.custom_xml_parts:
                values = cxml.xpath("//item/value/text()")
        """
        return self._part.custom_xml_parts

    def add_custom_xml(
        self,
        xml_content: str | bytes,
        namespace: str | None = None,
        root_tag: str = "root",
    ) -> "CustomXml":
        """Add a new custom XML part to the document.

        Args:
            xml_content: Either a complete XML string/bytes, or a dict-like
                        structure to be converted to XML.
            namespace: Optional namespace URI for the root element.
            root_tag: Tag name for the root element if creating from scratch.

        Returns:
            A CustomXml object providing access to the new custom XML part.

        Example::

            # Add from XML string
            custom_xml = document.add_custom_xml('''
                <customer>
                    <name>ACME Corp</name>
                    <contact>John Doe</contact>
                </customer>
            ''')

            # Access the item ID (for content control binding)
            print(f"Item ID: {custom_xml.item_id}")

            # Update values
            custom_xml.set_text("/customer/name", "New Corp")
        """
        from lxml import etree

        from docx.customxml import CustomXml

        # Parse or create the XML element
        if isinstance(xml_content, str):
            xml_content = xml_content.encode("utf-8")
        xml_element = etree.fromstring(xml_content)

        # Add to document
        xml_part, props_part = self._part.add_custom_xml(xml_element)

        return CustomXml(xml_part, props_part)

    @property
    def fields(self) -> Fields:
        """A |Fields| collection providing access to fields in this document.

        Fields are used for dynamic content such as page numbers, dates,
        cross-references, table of contents, citations, and more.

        This collection includes both simple fields (`<w:fldSimple>`) and
        complex fields (delimited by `<w:fldChar>` markers).

        Example::

            # Iterate over all fields
            for field in document.fields:
                print(f"{field.field_type}: {field.result}")

            # Filter by type
            for page_field in document.fields.filter_by_type("PAGE"):
                print(f"Page number: {page_field.result}")

            # Access simple vs complex fields separately
            print(f"Simple fields: {len(document.fields.simple)}")
            print(f"Complex fields: {len(document.fields.complex)}")

        Note:
            Field results are cached values from when the field was last
            updated. They may be stale if the document hasn't been opened
            in Word recently.
        """
        from docx.fields import Fields
        from docx.oxml.fields import iter_complex_fields

        # Get simple fields
        simple_fields = self._element.body.xpath(".//w:fldSimple")

        # Parse complex fields
        complex_fields = iter_complex_fields(self._element.body)

        return Fields(simple_fields, complex_fields)

    @property
    def endnotes(self) -> Endnotes:
        """A |Endnotes| collection providing access to endnotes in this document.

        Endnotes are notes that appear at the end of the document, referenced by
        superscript numbers or symbols in the document body.

        Example::

            # Iterate over all endnotes
            for endnote in document.endnotes:
                print(f"Endnote {endnote.endnote_id}: {endnote.text}")

            # Access a specific endnote by id
            endnote = document.endnotes[1]
            if endnote:
                print(endnote.text)
        """
        return self._part.endnotes

    @property
    def floating_shapes(self):
        """A |FloatingShapes| collection of floating/anchored shapes in this document.

        Floating shapes (also called anchored shapes) are positioned independently
        of the text flow and can have text wrap around them. This is in contrast
        to inline shapes which flow with the text.

        Example::

            # Iterate over floating shapes
            for shape in document.floating_shapes:
                print(f"Shape: {shape.name}, size: {shape.width}x{shape.height}")

            # Access by index
            if len(document.floating_shapes) > 0:
                first_shape = document.floating_shapes[0]
                print(f"Type: {first_shape.type}")

            # Check if behind text
            for shape in document.floating_shapes:
                if shape.is_behind_text:
                    print(f"{shape.name} is behind text")
        """
        from docx.shape import FloatingShapes

        return FloatingShapes(self._element.body, self._part)

    @property
    def footnotes(self) -> Footnotes:
        """A |Footnotes| collection providing access to footnotes in this document.

        Footnotes are notes that appear at the bottom of the page, referenced by
        superscript numbers or symbols in the document body.

        Example::

            # Iterate over all footnotes
            for footnote in document.footnotes:
                print(f"Footnote {footnote.footnote_id}: {footnote.text}")

            # Access a specific footnote by id
            footnote = document.footnotes[1]
            if footnote:
                print(footnote.text)

            # Get the number of footnotes
            print(f"Total footnotes: {len(document.footnotes)}")
        """
        return self._part.footnotes

    @property
    def inline_shapes(self):
        """The |InlineShapes| collection for this document.

        An inline shape is a graphical object, such as a picture, contained in a run of
        text and behaving like a character glyph, being flowed like other text in a
        paragraph.
        """
        return self._part.inline_shapes

    @property
    def math_equations(self) -> MathEquations:
        """A |MathEquations| collection of math equations in this document.

        Math equations use Office Math Markup Language (OMML). This collection
        provides read-only access to all math zones in the document.

        Example::

            # Check for math equations
            if document.math_equations:
                print(f"Found {len(document.math_equations)} equations")

            # Iterate over all equations
            for eq in document.math_equations:
                print(f"{'Block' if eq.is_block else 'Inline'}: {eq.text}")

            # Access inline vs block equations separately
            for eq in document.math_equations.inline:
                print(f"Inline: {eq.text}")
            for eq in document.math_equations.block:
                print(f"Block: {eq.text}")
        """
        from docx.math import MathEquations

        body = self._element.body

        # Find m:oMath elements that are NOT children of m:oMathPara
        # These are inline equations
        all_omath = body.xpath(".//m:oMath")
        math_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
        inline_omath = [
            elem for elem in all_omath
            if elem.getparent() is None or elem.getparent().tag != f"{math_ns}oMathPara"
        ]

        # Find m:oMathPara elements (block-level equations)
        omathpara = body.xpath(".//m:oMathPara")

        return MathEquations(inline_omath, omathpara)

    def iter_inner_content(self) -> Iterator[Paragraph | Table]:
        """Generate each `Paragraph` or `Table` in this document in document order."""
        return self._body.iter_inner_content()

    @property
    def paragraphs(self) -> List[Paragraph]:
        """The |Paragraph| instances in the document, in document order.

        Note that paragraphs within revision marks such as ``<w:ins>`` or ``<w:del>`` do
        not appear in this list.
        """
        return self._body.paragraphs

    @property
    def part(self) -> DocumentPart:
        """The |DocumentPart| object of this document."""
        return self._part

    @property
    def revisions(self) -> Revisions:
        """A |Revisions| collection providing read-only access to track changes.

        Track changes (revisions) record insertions, deletions, and formatting
        changes made to the document when revision tracking is enabled in Word.

        Example::

            # Check if document has revisions
            if len(document.revisions) > 0:
                print(f"Document has {len(document.revisions)} tracked changes")

            # Iterate over all revisions
            for revision in document.revisions:
                print(f"{revision.revision_type.value}: '{revision.text}' by {revision.author}")

            # Get only insertions or deletions
            insertions = document.revisions.insertions
            deletions = document.revisions.deletions

            # Get unique authors
            authors = document.revisions.authors

            # Filter by author
            for r in document.revisions.by_author("John Doe"):
                print(r.text)

        Note:
            This provides read-only access. Accepting or rejecting revisions
            is not currently supported.
        """
        from docx.revisions import Revisions

        insertions = self._element.body.xpath(".//w:ins")
        deletions = self._element.body.xpath(".//w:del")
        return Revisions(insertions, deletions)

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this document to `path_or_stream`.

        `path_or_stream` can be either a path to a filesystem location (a string) or a
        file-like object.
        """
        self._part.save(path_or_stream)

    @property
    def sections(self) -> Sections:
        """|Sections| object providing access to each section in this document."""
        return Sections(self._element, self._part)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the document-level settings."""
        return self._part.settings

    @property
    def smartart(self) -> SmartArtCollection:
        """A |SmartArtCollection| of SmartArt diagrams in this document.

        SmartArt diagrams are complex graphical objects that display information
        in organized layouts (hierarchies, cycles, lists, etc.).

        Example::

            # Check for SmartArt
            if document.smartart:
                print(f"Found {len(document.smartart)} SmartArt diagrams")

            # Iterate and extract text
            for diagram in document.smartart:
                print(f"Nodes: {diagram.node_count}")
                print(f"Text: {diagram.text}")

        Note:
            This provides detection and text extraction only. SmartArt layout
            and styling information is stored in separate parts.
        """
        from docx.smartart import SmartArtCollection

        body = self._element.body
        smartart_elements = []

        # Find all dgm:relIds elements (SmartArt references)
        dgm_rel_ids = body.xpath(".//dgm:relIds")

        # For each SmartArt reference, try to get the data model part
        for rel_ids in dgm_rel_ids:
            # Get the r:dm (data model) relationship ID
            dm_id = rel_ids.get(
                "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm"
            )
            if dm_id:
                try:
                    # Get the related data model part
                    dm_part = self._part.related_parts.get(dm_id)
                    if dm_part is not None:
                        smartart_elements.append(dm_part.element)
                except (KeyError, AttributeError):
                    pass

        return SmartArtCollection(smartart_elements)

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in this document."""
        return self._part.styles

    @property
    def theme(self):
        """A |Theme| object providing access to the document's theme.

        The theme defines the color scheme, font scheme, and effects used in the document.

        Example::

            # Access theme colors
            theme = document.theme
            print(f"Theme: {theme.name}")
            print(f"Accent 1: {theme.colors.accent1}")  # RGBColor

            # Access theme fonts
            print(f"Heading font: {theme.fonts.major_latin}")
            print(f"Body font: {theme.fonts.minor_latin}")
        """
        return self._part.theme

    @property
    def conformance(self) -> str:
        """Return the OOXML conformance class of this document.

        Returns either 'transitional' or 'strict'.

        Transitional conformance is the default format used by Word 2007 and later.
        Strict conformance uses different namespaces and is less common.

        Example::

            if document.conformance == 'strict':
                print("Warning: Strict conformance document - some features may not work")
        """
        # Check the main namespace of the document element
        nsmap = self._element.nsmap
        w_ns = nsmap.get("w", "")
        if "purl.oclc.org" in w_ns:
            return "strict"
        return "transitional"

    @property
    def minimum_word_version(self) -> str:
        """Return the minimum Microsoft Word version required to open this document.

        This is determined by analyzing the mc:Ignorable attribute which lists
        namespace prefixes for features that can be safely ignored by older versions.

        Returns one of:
        - 'Word 2007': No extended namespaces
        - 'Word 2010': Uses w14 namespace
        - 'Word 2013': Uses w15 namespace
        - 'Word 2016': Uses w16 namespace
        - 'Word 365/2019': Uses w16cid or w16cex
        - 'Word 2021+': Uses w16sdtdh

        Example::

            print(f"This document requires {document.minimum_word_version} or later")
        """
        # Get mc:Ignorable attribute from root element
        ignorable = self._element.get(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
        )
        if ignorable is None:
            return "Word 2007"

        prefixes = set(ignorable.split())

        # Check in order of newest to oldest
        if "w16sdtdh" in prefixes:
            return "Word 2021+"
        if "w16cex" in prefixes or "w16cid" in prefixes:
            return "Word 365/2019"
        if "w16" in prefixes or "w16se" in prefixes:
            return "Word 2016"
        if "w15" in prefixes:
            return "Word 2013"
        if "w14" in prefixes:
            return "Word 2010"
        return "Word 2007"

    @property
    def supported_namespaces(self) -> List[str]:
        """Return a list of Word extension namespace prefixes used in this document.

        These are extracted from the mc:Ignorable attribute on the document element.
        Common prefixes include:
        - w14: Word 2010 features
        - w15: Word 2013 features
        - w16, w16se, w16cid, w16cex: Word 2016/365 features
        - w16sdtdh: Word 2021+ features
        - wp14: Word 2010 drawing features

        Example::

            namespaces = document.supported_namespaces
            print(f"This document uses: {', '.join(namespaces)}")

            if 'w15' in namespaces:
                print("Document uses Word 2013 features")
        """
        ignorable = self._element.get(
            "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
        )
        if ignorable is None:
            return []
        return ignorable.split()

    @property
    def tables(self) -> List[Table]:
        """All |Table| instances in the document, in document order.

        Note that only tables appearing at the top level of the document appear in this
        list; a table nested inside a table cell does not appear. A table within
        revision marks such as ``<w:ins>`` or ``<w:del>`` will also not appear in the
        list.
        """
        return self._body.tables

    @property
    def text_boxes(self) -> TextBoxes:
        """A |TextBoxes| collection providing access to text boxes in this document.

        Text boxes are floating containers that can hold paragraphs and tables.
        They are stored inside mc:AlternateContent elements in the document.

        Example::

            # Iterate over all text boxes
            for textbox in document.text_boxes:
                print(f"Text box with {len(textbox.paragraphs)} paragraphs")
                for para in textbox.paragraphs:
                    print(f"  {para.text}")

            # Access by index
            if document.text_boxes:
                first_box = document.text_boxes[0]
                print(first_box.text)

            # Modify text box content
            document.text_boxes[0].paragraphs[0].text = "Updated text"
        """
        from docx.textbox import TextBoxes

        alt_contents = self._element.body.xpath(".//mc:AlternateContent")
        return TextBoxes(alt_contents, self._part)

    @property
    def _block_width(self) -> Length:
        """A |Length| object specifying the space between margins in last section."""
        section = self.sections[-1]
        page_width = section.page_width or Inches(8.5)
        left_margin = section.left_margin or Inches(1)
        right_margin = section.right_margin or Inches(1)
        return Emu(page_width - left_margin - right_margin)

    @property
    def _body(self) -> _Body:
        """The |_Body| instance containing the content for this document."""
        if self.__body is None:
            self.__body = _Body(self._element.body, self)
        return self.__body


class _Body(BlockItemContainer):
    """Proxy for `<w:body>` element in this document.

    It's primary role is a container for document content.
    """

    def __init__(self, body_elm: CT_Body, parent: t.ProvidesStoryPart):
        super(_Body, self).__init__(body_elm, parent)
        self._body = body_elm

    def clear_content(self) -> _Body:
        """Return this |_Body| instance after clearing it of all content.

        Section properties for the main document story, if present, are preserved.
        """
        self._body.clear_content()
        return self
