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
    from docx.bookmarks import Bookmarks
    from docx.comments import Comment, Comments
    from docx.fields import Fields
    from docx.footnotes import Endnotes, Footnotes
    from docx.oxml.document import CT_Body, CT_Document
    from docx.parts.document import DocumentPart
    from docx.revisions import Revisions
    from docx.sdt import SdtContentControls
    from docx.settings import Settings
    from docx.styles.style import ParagraphStyle, _TableStyle
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    from docx.textbox import TextBoxes


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

    @property
    def core_properties(self):
        """A |CoreProperties| object providing Dublin Core properties of document."""
        return self._part.core_properties

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
