# python-docx API Reference

This document provides a comprehensive overview of the python-docx API, including both the upstream features and WOTAN extensions.

---

## Table of Contents

1. [Getting Started](#getting-started)
2. [Document](#document)
3. [Paragraphs and Runs](#paragraphs-and-runs)
4. [Tables](#tables)
5. [Sections and Page Layout](#sections-and-page-layout)
6. [Styles](#styles)
7. [Images and Shapes](#images-and-shapes)
8. [Headers and Footers](#headers-and-footers)
9. [Fields](#fields)
10. [Bookmarks](#bookmarks)
11. [Content Controls (SDT)](#content-controls-sdt)
12. [Footnotes and Endnotes](#footnotes-and-endnotes)
13. [Comments](#comments)
14. [Track Changes (Revisions)](#track-changes-revisions)
15. [Text Boxes](#text-boxes)
16. [Theme](#theme)
17. [Document Properties](#document-properties)
18. [Pragmatics: Real-World Document Patterns](#pragmatics-real-world-document-patterns)

---

## Getting Started

```python
from docx import Document
from docx.shared import Inches, Pt, Cm, Emu, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT

# Create new document
doc = Document()

# Open existing document
doc = Document('existing.docx')

# Save document
doc.save('output.docx')
```

### Length Units

| Class | Description | Example |
|-------|-------------|---------|
| `Inches(n)` | Inches | `Inches(1.5)` |
| `Pt(n)` | Points (1/72 inch) | `Pt(12)` |
| `Cm(n)` | Centimeters | `Cm(2.54)` |
| `Emu(n)` | English Metric Units (914400 per inch) | `Emu(914400)` |
| `Twips(n)` | Twips (1/20 point) | `Twips(240)` |

---

## Document

The `Document` object is the root of the document tree.

### Creating Content

```python
# Add paragraph
para = doc.add_paragraph('Text content')
para = doc.add_paragraph('Styled text', style='Heading 1')

# Add heading (level 0-9, where 0 is Title)
heading = doc.add_heading('Chapter Title', level=1)

# Add page break
doc.add_page_break()

# Add table
table = doc.add_table(rows=3, cols=4)
table = doc.add_table(rows=3, cols=4, style='Table Grid')

# Add picture (inline)
doc.add_picture('image.png', width=Inches(2))

# Add floating picture (WOTAN extension)
doc.add_floating_picture(
    'image.png',
    width=Inches(2),
    pos_x=Inches(1),
    pos_y=Inches(2),
    wrap_type='square',      # none, square, tight, through, topAndBottom
    behind_doc=False,
    h_relative_from='column',
    v_relative_from='paragraph'
)

# Add section break
section = doc.add_section(start_type=WD_SECTION.NEW_PAGE)

# Add bookmark (WOTAN extension)
bookmark = doc.add_bookmark('my_bookmark', start=paragraph)

# Add content control (WOTAN extension)
cc = doc.add_content_control(
    sdt_type='richText',  # richText, text, date, dropDownList, comboBox
    tag='field_name',
    alias='Display Name'
)

# Add comment (requires runs to anchor to)
comment = doc.add_comment(runs=para.runs, text='Comment text', author='Name')
```

### Accessing Content

```python
# Paragraphs (list)
for para in doc.paragraphs:
    print(para.text)

# Tables (list)
for table in doc.tables:
    for row in table.rows:
        for cell in row.cells:
            print(cell.text)

# Iterate all block content in order
for block in doc.iter_inner_content():
    if hasattr(block, 'text'):  # Paragraph
        print(block.text)
    else:  # Table
        print(f"Table: {len(block.rows)} rows")

# Sections
for section in doc.sections:
    print(section.page_width, section.page_height)

# Styles
for style in doc.styles:
    print(style.name, style.type)
```

### Document Metadata (WOTAN Extensions)

```python
# OOXML conformance class
doc.conformance          # 'transitional' or 'strict'

# Minimum Word version required
doc.minimum_word_version # 'Word 2007', 'Word 2010', ..., 'Word 2021+'

# Extension namespaces used
doc.supported_namespaces # ['w14', 'w15', 'w16', ...]
```

---

## Paragraphs and Runs

### Paragraph

```python
para = doc.add_paragraph()

# Text content
para.text                    # Get full text
para.text = 'New text'       # Replace all content

# Style
para.style = 'Heading 1'
para.style = doc.styles['Normal']

# Alignment
para.alignment = WD_ALIGN_PARAGRAPH.CENTER  # LEFT, CENTER, RIGHT, JUSTIFY

# Formatting (via paragraph_format)
fmt = para.paragraph_format
fmt.left_indent = Inches(0.5)
fmt.right_indent = Inches(0.5)
fmt.first_line_indent = Inches(0.25)
fmt.space_before = Pt(12)
fmt.space_after = Pt(12)
fmt.line_spacing = 1.5
fmt.keep_together = True
fmt.keep_with_next = True
fmt.page_break_before = True
fmt.widow_control = True

# Runs
for run in para.runs:
    print(run.text)

# Add run
run = para.add_run('text')
run = para.add_run('bold text').bold = True

# Clear paragraph
para.clear()

# Add bookmark to paragraph (WOTAN extension)
bookmark = para.add_bookmark('bookmark_name')
```

### Run

```python
run = para.add_run('text')

# Text
run.text = 'new text'

# Character formatting (via font)
font = run.font
font.name = 'Arial'
font.size = Pt(12)
font.bold = True
font.italic = True
font.underline = True                    # or WD_UNDERLINE.SINGLE, DOUBLE, etc.
font.strike = True
font.double_strike = True
font.subscript = True
font.superscript = True
font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
font.color.theme_color = MSO_THEME_COLOR.ACCENT_1
font.highlight_color = WD_COLOR_INDEX.YELLOW
font.all_caps = True
font.small_caps = True

# Add content to run
run.add_break()                          # Line break
run.add_break(WD_BREAK.PAGE)             # Page break
run.add_break(WD_BREAK.COLUMN)           # Column break
run.add_tab()
run.add_picture('image.png', width=Inches(1))

# Add footnote/endnote reference (WOTAN extension)
footnote = run.add_footnote_reference()
footnote.add_paragraph('Footnote text')

endnote = run.add_endnote_reference()
endnote.add_paragraph('Endnote text')

# Add bookmark (WOTAN extension)
bookmark = run.add_bookmark('bookmark_name')

# Add field (WOTAN extension)
run.add_simple_field('PAGE')
run.add_complex_field('TOC \\o "1-3"')
```

---

## Tables

```python
table = doc.add_table(rows=3, cols=4)

# Style
table.style = 'Table Grid'

# Alignment
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Access cells
cell = table.cell(0, 0)          # By row, col index
cell = table.rows[0].cells[0]    # Via row
cell = table.columns[0].cells[0] # Via column

# Cell content
cell.text = 'Content'
cell.paragraphs                  # List of paragraphs in cell
cell.add_paragraph('More text')
cell.add_table(rows=2, cols=2)   # Nested table

# Cell formatting
cell.width = Inches(2)
cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Merge cells
cell_a = table.cell(0, 0)
cell_b = table.cell(0, 2)
cell_a.merge(cell_b)             # Merge horizontally

# Row operations
row = table.rows[0]
row.height = Inches(0.5)
row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

# Add row
table.add_row()

# Column width
table.columns[0].width = Inches(1.5)
```

---

## Sections and Page Layout

```python
section = doc.sections[0]        # First section
section = doc.sections[-1]       # Last section
section = doc.add_section()      # Add new section

# Page size
section.page_width = Inches(8.5)
section.page_height = Inches(11)

# Orientation
section.orientation = WD_ORIENT.LANDSCAPE  # or PORTRAIT

# Margins
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1.25)
section.right_margin = Inches(1.25)
section.gutter = Inches(0)

# Headers/Footers distance
section.header_distance = Inches(0.5)
section.footer_distance = Inches(0.5)

# Columns (not fully supported)
# section.cols  # Limited support

# Section break type
section.start_type = WD_SECTION.NEW_PAGE  # CONTINUOUS, NEW_COLUMN, etc.

# Different first page header/footer
section.different_first_page_header_footer = True
```

---

## Styles

```python
# Access styles
styles = doc.styles

# Get specific style
normal = styles['Normal']
heading1 = styles['Heading 1']

# Style properties
style = styles['Normal']
style.name                    # Style name
style.type                    # WD_STYLE_TYPE.PARAGRAPH, CHARACTER, TABLE, LIST
style.base_style              # Parent style
style.font                    # Font formatting
style.paragraph_format        # Paragraph formatting (for paragraph styles)

# Create new style
new_style = styles.add_style('My Style', WD_STYLE_TYPE.PARAGRAPH)
new_style.base_style = styles['Normal']
new_style.font.bold = True
new_style.paragraph_format.space_after = Pt(12)

# Character style
char_style = styles.add_style('Emphasis', WD_STYLE_TYPE.CHARACTER)
char_style.font.italic = True
```

---

## Images and Shapes

### Inline Shapes

```python
# Add inline picture
run.add_picture('image.png')
run.add_picture('image.png', width=Inches(2))
run.add_picture('image.png', height=Inches(3))
run.add_picture('image.png', width=Inches(2), height=Inches(1.5))

# Access inline shapes
for shape in doc.inline_shapes:
    print(shape.width, shape.height)
    print(shape.type)  # PICTURE, CHART, etc.
```

### Floating Shapes (WOTAN Extension)

```python
# Add floating picture
shape = doc.add_floating_picture(
    'image.png',
    width=Inches(2),
    height=Inches(1.5),
    pos_x=Inches(1),
    pos_y=Inches(2),
    behind_doc=False,
    wrap_type='square',        # none, square, tight, through, topAndBottom
    h_relative_from='column',  # character, column, margin, page, etc.
    v_relative_from='paragraph'
)

# Access floating shapes
for shape in doc.floating_shapes:
    print(shape.name)
    print(shape.width, shape.height)
    print(shape.type)
    print(shape.is_behind_text)
```

---

## Headers and Footers

```python
section = doc.sections[0]

# Access header/footer
header = section.header
footer = section.footer

# First page header/footer (if different)
section.different_first_page_header_footer = True
first_header = section.first_page_header
first_footer = section.first_page_footer

# Even page header/footer (if different odd/even)
# section.different_odd_and_even_pages = True  # Limited support
# even_header = section.even_page_header

# Header/footer content
header.paragraphs[0].text = 'Header text'
header.add_paragraph('More header text')

# Link to previous section
header.is_linked_to_previous = False

# Check if header has content
if header.is_linked_to_previous:
    print("Using previous section's header")
```

---

## Fields (WOTAN Extension)

Fields are dynamic content placeholders that Word updates.

```python
# Access all fields
fields = doc.fields
print(f"Total fields: {len(fields)}")

# Simple vs complex fields
simple_fields = fields.simple    # List of SimpleField
complex_fields = fields.complex  # List of ComplexFieldProxy

# Iterate all fields
for field in doc.fields:
    print(field.field_type)    # PAGE, DATE, TOC, REF, HYPERLINK, etc.
    print(field.field_code)    # Full field code
    print(field.result)        # Cached result text

# Filter by type
for field in fields.filter_by_type('PAGE'):
    print(f"Page number: {field.result}")

# Create fields (on a run)
run.add_simple_field('PAGE')
run.add_simple_field('DATE \\@ "yyyy-MM-dd"')
run.add_complex_field('TOC \\o "1-3" \\h')
```

### Common Field Types

| Type | Description | Example Code |
|------|-------------|--------------|
| `PAGE` | Current page number | `PAGE` |
| `NUMPAGES` | Total pages | `NUMPAGES` |
| `DATE` | Current date | `DATE \@ "yyyy-MM-dd"` |
| `TIME` | Current time | `TIME \@ "HH:mm"` |
| `TOC` | Table of contents | `TOC \o "1-3" \h` |
| `REF` | Cross-reference | `REF bookmark_name` |
| `PAGEREF` | Page of bookmark | `PAGEREF bookmark_name` |
| `HYPERLINK` | Hyperlink | `HYPERLINK "url"` |
| `SEQ` | Sequence number | `SEQ Figure` |
| `STYLEREF` | Style reference | `STYLEREF "Heading 1"` |
| `CITATION` | Bibliography citation | `CITATION key` |

---

## Bookmarks (WOTAN Extension)

Bookmarks are named locations in the document.

```python
# Access bookmarks
bookmarks = doc.bookmarks
print(f"Total bookmarks: {len(bookmarks)}")

# Check if bookmark exists
if 'my_bookmark' in bookmarks:
    print("Found!")

# Get bookmark by name
bookmark = bookmarks.get('my_bookmark')
if bookmark:
    print(bookmark.name)
    print(bookmark.bookmark_id)

# Iterate bookmarks
for bookmark in doc.bookmarks:
    print(f"{bookmark.name}: id={bookmark.bookmark_id}")

# Filter user vs system bookmarks
user_bookmarks = [bm for bm in bookmarks if not bm.is_system]
system_bookmarks = [bm for bm in bookmarks if bm.is_system]

# System bookmark types
bookmark.is_toc_bookmark      # _Toc...
bookmark.is_ref_bookmark      # _Ref...
bookmark.is_hyperlink_bookmark # _Hlt...

# Create bookmark
bookmark = doc.add_bookmark('new_bookmark', start=paragraph)
bookmark = para.add_bookmark('para_bookmark')
bookmark = run.add_bookmark('run_bookmark')
```

---

## Content Controls (SDT) (WOTAN Extension)

Content controls (Structured Document Tags) are used for forms and data binding.

```python
# Access content controls
ccs = doc.content_controls
print(f"Total content controls: {len(ccs)}")

# Iterate
for cc in doc.content_controls:
    print(cc.sdt_type)   # richText, text, date, dropDownList, comboBox, docPartObj
    print(cc.tag)        # Programmatic identifier
    print(cc.alias)      # Display name
    print(cc.text)       # Text content

# Find by tag
cc = ccs.get_by_tag('customer_name')

# Find by alias
cc = ccs.get_by_alias('Customer Name')

# Content control is a BlockItemContainer
for para in cc.paragraphs:
    print(para.text)

# Create content control
cc = doc.add_content_control(
    sdt_type='dropDownList',
    tag='status',
    alias='Status'
)
cc.add_list_item('Draft', 'draft')
cc.add_list_item('Final', 'final')
```

---

## Footnotes and Endnotes (WOTAN Extension)

```python
# Access footnotes/endnotes
footnotes = doc.footnotes
endnotes = doc.endnotes

print(f"Footnotes: {len(footnotes)}")
print(f"Endnotes: {len(endnotes)}")

# Iterate
for fn in doc.footnotes:
    print(fn.footnote_id)
    print(fn.text)
    # Footnote is a BlockItemContainer
    for para in fn.paragraphs:
        print(para.text)

# Access by ID
footnote = footnotes[1]  # Get footnote with id=1

# Create footnote (from a run)
run = para.add_run('Text with footnote')
footnote = run.add_footnote_reference()
footnote.add_paragraph('This is the footnote text.')

# Create endnote
endnote = run.add_endnote_reference()
endnote.add_paragraph('This is the endnote text.')
```

---

## Comments (WOTAN Extension)

```python
# Access comments
comments = doc.comments
print(f"Total comments: {len(comments)}")

# Iterate
for comment in doc.comments:
    print(comment.author)
    print(comment.initials)
    print(comment.date)
    print(comment.text)
    print(comment.comment_id)

# Add comment
comment = doc.add_comment(
    runs=para.runs,        # Single run or sequence of runs
    text='Comment text',
    author='Author Name',
    initials='AN'
)

# Comment is a BlockItemContainer
comment.add_paragraph('Additional paragraph in comment')
```

---

## Track Changes (Revisions) (WOTAN Extension)

```python
# Access revisions
revisions = doc.revisions
print(f"Total revisions: {len(revisions)}")

# Separate insertions and deletions
insertions = revisions.insertions
deletions = revisions.deletions

# Get authors
authors = revisions.authors  # Set of author names

# Filter by author
for rev in revisions.by_author('John Doe'):
    print(rev.text)

# Iterate
for rev in doc.revisions:
    print(rev.revision_type)  # RevisionType.INSERTION or DELETION
    print(rev.author)
    print(rev.date)
    print(rev.text)

# Accept/reject individual revisions
rev.accept()   # Apply the change
rev.reject()   # Revert the change

# Bulk accept/reject
revisions.accept_all()
revisions.reject_all()
```

---

## Text Boxes (WOTAN Extension)

Text boxes are floating containers for content, stored in `mc:AlternateContent` elements.

```python
# Access text boxes
text_boxes = doc.text_boxes
print(f"Total text boxes: {len(text_boxes)}")

# Iterate
for tb in doc.text_boxes:
    print(tb.text)
    # TextBox is a BlockItemContainer
    for para in tb.paragraphs:
        print(para.text)
    for table in tb.tables:
        print(f"Table: {len(table.rows)} rows")

# Access by index
if text_boxes:
    first_box = text_boxes[0]
    print(first_box.paragraphs[0].text)

# Modify content
text_boxes[0].paragraphs[0].text = 'New text'
```

---

## Theme (WOTAN Extension)

The theme defines colors and fonts for the document.

```python
theme = doc.theme

# Theme name
print(theme.name)  # e.g., "Office Theme"

# Theme colors
colors = theme.colors
print(colors.dark1)        # RGBColor - main text
print(colors.light1)       # RGBColor - background
print(colors.dark2)        # RGBColor - secondary text
print(colors.light2)       # RGBColor - secondary background
print(colors.accent1)      # RGBColor - accent colors 1-6
print(colors.accent2)
print(colors.accent3)
print(colors.accent4)
print(colors.accent5)
print(colors.accent6)
print(colors.hyperlink)
print(colors.followed_hyperlink)

# Modify theme colors
colors.accent1 = RGBColor(0xFF, 0x00, 0x00)  # Red
colors.accent2 = '00FF00'                     # Green (hex string)
colors.hyperlink = '#0000FF'                  # Blue (with #)

# Theme fonts
fonts = theme.fonts
print(fonts.major_latin)         # Heading font
print(fonts.minor_latin)         # Body font
print(fonts.major_east_asian)
print(fonts.minor_east_asian)
print(fonts.major_complex_script)
print(fonts.minor_complex_script)

# Modify theme fonts
fonts.major_latin = 'Arial'
fonts.minor_latin = 'Times New Roman'
```

---

## Document Properties

```python
# Core properties (Dublin Core)
props = doc.core_properties

# Read properties
print(props.title)
print(props.author)
print(props.subject)
print(props.keywords)
print(props.category)
print(props.comments)
print(props.created)          # datetime
print(props.modified)         # datetime
print(props.last_modified_by)
print(props.revision)         # int
print(props.version)

# Modify properties
props.title = 'Document Title'
props.author = 'Author Name'
props.subject = 'Subject'
props.keywords = 'keyword1, keyword2'
props.category = 'Category'
props.comments = 'Description'
```

---

## Pragmatics: Real-World Document Patterns

This section documents practical patterns and gotchas discovered when working with real-world Word documents.

### 1. Text May Be in Unexpected Places

**Issue:** When searching for text visually present in a document, it may not be in `doc.paragraphs`.

**Common locations for "hidden" text:**

| Location | How to Access | Example |
|----------|---------------|---------|
| Headers/Footers | `section.header.paragraphs` | Page numbers, document titles |
| Text Boxes | `doc.text_boxes` | Callouts, sidebars |
| Content Controls | `doc.content_controls` | Form fields, template placeholders |
| Table Cells | `table.cell(r,c).paragraphs` | Structured data |
| Footnotes/Endnotes | `doc.footnotes`, `doc.endnotes` | References |
| Comments | `doc.comments` | Review annotations |
| Field Results | `doc.fields[n].result` | Dynamic content (dates, page numbers) |

**Example: Finding all text in a document:**

```python
def get_all_text(doc):
    """Extract text from all document locations."""
    text_parts = []

    # Main body
    for para in doc.paragraphs:
        text_parts.append(para.text)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text_parts.append(cell.text)

    # Headers and footers
    for section in doc.sections:
        for para in section.header.paragraphs:
            text_parts.append(para.text)
        for para in section.footer.paragraphs:
            text_parts.append(para.text)

    # Text boxes
    for tb in doc.text_boxes:
        text_parts.append(tb.text)

    # Content controls
    for cc in doc.content_controls:
        text_parts.append(cc.text)

    # Footnotes
    for fn in doc.footnotes:
        text_parts.append(fn.text)

    return '\n'.join(text_parts)
```

### 2. Nested Content Controls

**Issue:** Content controls (SDTs) can be nested inside other content controls. The `doc.content_controls` property returns top-level SDTs, but text may be in nested ones.

**Pattern observed:** Professional templates (like RISE reports) use nested SDTs for title page metadata:

```
document -> body -> sdt (docPartObj) -> sdtContent -> sdt (text) -> sdtContent -> p
```

**Solution:** Access the `text` property of a content control, which includes all nested text, or iterate nested SDTs via the XML layer.

### 3. Template Structures

**Common template patterns:**

| Pattern | Structure | Example |
|---------|-----------|---------|
| Title page metadata | Nested SDTs in first top-level SDT | Report number, author, date |
| Table of Contents | SDT with `docPartObj` type | Auto-updated TOC |
| Form fields | SDTs with `text`, `date`, `dropDownList` types | Fillable forms |
| Building blocks | SDT with `docPartObj` type | Cover pages, headers |

### 4. Machine-Generated Documents

**Issue:** Documents from other tools (Teams transcripts, PDF converters, mail merge) may have unusual structures.

**Patterns:**

| Source | Characteristics |
|--------|-----------------|
| Teams transcripts | No theme, all paragraphs unstyled, many floating shapes (avatars) |
| PDF to DOCX | Often uses positioned text boxes instead of paragraphs |
| Mail merge | Extensive use of fields and bookmarks |
| Scanned documents | Single image per page, no real text |

### 5. Field Codes vs Results

**Issue:** Field results are cached values that may be stale.

```python
# Field structure:
#   field_code: The instruction (e.g., "PAGE")
#   result: The cached value (e.g., "5")

# The result is what was displayed when the document was last opened in Word
# It may not reflect current state

for field in doc.fields:
    print(f"{field.field_type}: '{field.result}' (code: {field.field_code})")
```

### 6. Style Inheritance

**Issue:** Formatting may come from styles, not direct formatting.

```python
# A paragraph with no direct formatting inherits from its style
para.style = 'Heading 1'  # Font, size, spacing from style

# Direct formatting overrides style
para.runs[0].font.bold = True  # This run is bold regardless of style

# To find effective formatting, check both
effective_bold = run.font.bold
if effective_bold is None:
    # Check paragraph style, then character style
    pass
```

### 7. Document Size and Performance

**Observations from real documents:**

| Document Size | Load Time | Characteristics |
|---------------|-----------|-----------------|
| < 1 MB | < 1 sec | Typical text documents |
| 1-10 MB | 1-5 sec | Documents with images |
| 10-50 MB | 5-15 sec | Many images, complex formatting |
| 50-100+ MB | 15-60 sec | Large reports with embedded media |

**The 98.8 MB nautical analysis document loaded in ~15 seconds** and had:
- 86 inline images
- 433 complex fields
- 347 bookmarks
- 27 tables

### 8. Empty Paragraphs and Whitespace

**Issue:** Documents often have many empty paragraphs for spacing.

```python
# The nautical analysis document had 12 empty paragraphs before content
# This is common for title page vertical spacing

# Filter non-empty paragraphs
content_paras = [p for p in doc.paragraphs if p.text.strip()]
```

### 9. Cross-References and Bookmarks

**Pattern:** Complex documents use extensive cross-referencing:

```
SEQ fields -> Number figures/tables
REF fields -> Reference those numbers
PAGEREF fields -> Reference page numbers
Bookmarks -> Mark reference targets
```

**Example from nautical analysis:**
- 118 REF fields
- 105 STYLEREF fields
- 105 SEQ fields
- 60 PAGEREF fields
- 347 system bookmarks

### 10. Track Changes Workflow

**Common pattern:** Documents with track changes need cleanup before final distribution.

```python
# Check for pending changes
if len(doc.revisions) > 0:
    print(f"Document has {len(doc.revisions)} tracked changes")
    print(f"Authors: {doc.revisions.authors}")

    # Accept all before distribution
    doc.revisions.accept_all()
    doc.save('final.docx')
```

---

## Version Information

| Feature | Availability |
|---------|--------------|
| Core API (paragraphs, tables, styles) | python-docx upstream |
| Fields (read) | WOTAN extension |
| Fields (create) | WOTAN extension |
| Bookmarks (read) | WOTAN extension |
| Bookmarks (create) | WOTAN extension |
| Content Controls (read) | WOTAN extension |
| Content Controls (create) | WOTAN extension |
| Footnotes/Endnotes (read) | WOTAN extension |
| Footnotes/Endnotes (create) | WOTAN extension |
| Comments | python-docx upstream + WOTAN |
| Track Changes (read) | WOTAN extension |
| Track Changes (accept/reject) | WOTAN extension |
| Text Boxes | WOTAN extension |
| Floating Shapes (read) | WOTAN extension |
| Floating Shapes (create) | WOTAN extension |
| Theme (read) | WOTAN extension |
| Theme (modify) | WOTAN extension |
| Conformance Detection | WOTAN extension |
