# Current State: What Works and What Doesn't

**Last updated:** 2025-12-03 (after completing all WOTAN extensions)

This document provides a quick reference for what python-docx can and cannot do.

## Executive Summary

python-docx 1.2.0 with WOTAN extensions handles **nearly all document operations**. Basic content (paragraphs, tables, styles) has full support. Advanced content (fields, content controls, footnotes, bookmarks, track changes, floating shapes, text boxes, themes, math, charts, SmartArt, custom XML) is now accessible via the extended API.

**Coverage estimate:** ~95% for document content access and manipulation.

---

## What Works (Full API Access)

### Document Structure
```python
doc = Document('file.docx')
doc.paragraphs          # ✅ All paragraphs
doc.tables              # ✅ All tables
doc.sections            # ✅ All sections
doc.styles              # ✅ All styles
doc.inline_shapes       # ✅ Inline images
doc.comments            # ✅ Comments (v1.2.0)
doc.core_properties     # ✅ Title, author, etc.
```

### WOTAN Extensions
```python
# All these are now fully accessible:
doc.content_controls    # ✅ SDT access with full CRUD
doc.fields              # ✅ Simple and complex fields
doc.bookmarks           # ✅ With rename/delete support
doc.footnotes           # ✅ Full content access
doc.endnotes            # ✅ Full content access
doc.revisions           # ✅ With accept/reject
doc.floating_shapes     # ✅ With full modification
doc.text_boxes          # ✅ With creation support
doc.theme               # ✅ Colors and fonts, read/modify
doc.equations           # ✅ Math equations, read/create
doc.charts              # ✅ Detection and access
doc.smartart_objects    # ✅ Detection and access
doc.custom_xml_parts    # ✅ Full CRUD
doc.conformance         # ✅ Strict vs Transitional
doc.minimum_word_version # ✅ Version detection
```

### Paragraph & Run Access
```python
for p in doc.paragraphs:
    p.text                    # ✅ Get/set text
    p.style                   # ✅ Get/set style
    p.alignment               # ✅ Get/set alignment
    p.paragraph_format        # ✅ Spacing, indentation, etc.

    for run in p.runs:
        run.text              # ✅ Get/set text
        run.bold              # ✅ Get/set bold
        run.italic            # ✅ Get/set italic
        run.font.name         # ✅ Font name
        run.font.size         # ✅ Font size
        run.font.color.rgb    # ✅ Font color
        # ... all character formatting
```

### Tables
```python
for table in doc.tables:
    table.rows                # ✅ All rows
    table.columns             # ✅ All columns
    for row in table.rows:
        for cell in row.cells:
            cell.text         # ✅ Cell text
            cell.paragraphs   # ✅ Cell paragraphs
            cell.tables       # ✅ Nested tables
```

### Sections & Headers/Footers
```python
for section in doc.sections:
    section.page_width        # ✅ Page dimensions
    section.page_height
    section.top_margin        # ✅ Margins
    section.header            # ✅ Header access
    section.footer            # ✅ Footer access
    section.different_first_page_header_footer  # ✅
```

### Images
```python
for shape in doc.inline_shapes:
    shape.type                # ✅ PICTURE, CHART, etc.
    shape.width               # ✅ Get/set dimensions
    shape.height

# Add new inline image
doc.add_picture('image.png', width=Inches(2))

# Floating images (WOTAN extension)
doc.add_floating_picture('image.png', width=Inches(2), pos_x=Inches(1), pos_y=Inches(2))

for shape in doc.floating_shapes:
    shape.width = Inches(3)   # ✅ Resize
    shape.pos_x = Inches(1)   # ✅ Reposition
    shape.name = "NewName"    # ✅ Rename
    shape.delete()            # ✅ Remove
```

### Fields (WOTAN Extension)
```python
for field in doc.fields:
    field.field_type          # ✅ PAGE, DATE, TOC, REF, etc.
    field.field_code          # ✅ Full instruction
    field.result              # ✅ Cached value
    field.delete()            # ✅ Remove from document
    field.convert_to_text()   # ✅ Convert to static text

# Create fields
run.add_simple_field('PAGE')
run.add_complex_field('TOC \\o "1-3"')
```

### Content Controls (WOTAN Extension)
```python
for cc in doc.content_controls:
    cc.sdt_type               # ✅ richText, text, date, dropDownList, etc.
    cc.tag                    # ✅ Programmatic identifier
    cc.alias                  # ✅ Display name
    cc.text                   # ✅ Get/set content

# Create content controls
doc.add_content_control(sdt_type='text', tag='field_name')
```

### Bookmarks (WOTAN Extension)
```python
for bookmark in doc.bookmarks:
    bookmark.name             # ✅ Get/set name
    bookmark.bookmark_id      # ✅ Internal ID
    bookmark.delete()         # ✅ Remove from document

# Create bookmarks
doc.add_bookmark('my_bookmark', start=paragraph)
```

### Track Changes (WOTAN Extension)
```python
for rev in doc.revisions:
    rev.revision_type         # ✅ INSERTION or DELETION
    rev.author                # ✅ Who made the change
    rev.date                  # ✅ When
    rev.text                  # ✅ Changed content
    rev.accept()              # ✅ Apply change
    rev.reject()              # ✅ Revert change

doc.revisions.accept_all()    # ✅ Bulk accept
doc.revisions.reject_all()    # ✅ Bulk reject
```

### Text Boxes (WOTAN Extension)
```python
for tb in doc.text_boxes:
    tb.paragraphs             # ✅ Access content
    tb.tables                 # ✅ Tables in text box

# Create text boxes
text_box = doc.add_text_box(width=Inches(2), height=Inches(1))
text_box.paragraphs[0].text = "Hello!"
```

### Math Equations (WOTAN Extension)
```python
for eq in doc.equations:
    eq.latex                  # ✅ LaTeX representation (if available)

# Create equations
para.add_equation('x^2 + y^2 = z^2')
```

### Theme (WOTAN Extension)
```python
theme = doc.theme
theme.colors.accent1          # ✅ Read theme colors
theme.colors.accent1 = RGBColor(0xFF, 0x00, 0x00)  # ✅ Modify
theme.fonts.major_latin       # ✅ Read theme fonts
theme.fonts.major_latin = 'Arial'  # ✅ Modify
```

---

## What Has Limited Support

### Charts — Detection Only
```python
for chart in doc.charts:
    chart.name                # ✅ Access name
    # chart.data              # ❌ Cannot modify chart data
```

### SmartArt — Detection Only
```python
for smartart in doc.smartart_objects:
    smartart.name             # ✅ Access name
    # Cannot modify layout or content programmatically
```

### Numbering Definitions — ⚠️ PARTIAL
```python
# Can apply existing numbering:
p.style = 'List Bullet'       # ✅ Works

# Cannot create new numbering definitions:
doc.numbering.add_definition()  # ❌ Not implemented
```

---

## Preservation Behavior

Unsupported elements **survive round-trip**:

```python
doc = Document('complex.docx')  # Has any content...
doc.add_paragraph('New text')   # Make a change
doc.save('output.docx')         # Save

# Result: output.docx still has all original content intact!
```

---

## Quick Reference Card

| Feature | Read | Modify | Create |
|---------|:----:|:------:|:------:|
| Paragraphs | ✅ | ✅ | ✅ |
| Runs/Text | ✅ | ✅ | ✅ |
| Formatting | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | ✅ |
| Styles | ✅ | ✅ | ✅ |
| Sections | ✅ | ✅ | ✅ |
| Headers/Footers | ✅ | ✅ | ✅ |
| Inline Images | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | ✅ |
| Comments | ✅ | ✅ | ✅ |
| Core Properties | ✅ | ✅ | ✅ |
| **Fields** | ✅ | ✅ | ✅ |
| **Content Controls** | ✅ | ✅ | ✅ |
| **Bookmarks** | ✅ | ✅ | ✅ |
| **Footnotes/Endnotes** | ✅ | ✅ | ✅ |
| **Track Changes** | ✅ | ✅ | — |
| **Floating Shapes** | ✅ | ✅ | ✅ |
| **Text Boxes** | ✅ | ✅ | ✅ |
| **Theme** | ✅ | ✅ | — |
| **Math Equations** | ✅ | ✅ | ✅ |
| **Charts** | ✅ | — | — |
| **SmartArt** | ✅ | — | — |
| **Custom XML** | ✅ | ✅ | ✅ |
| **Numbering (new)** | ✅ | ⚠️ | — |

**Legend:** ✅ Full | ⚠️ Partial | — Not implemented
