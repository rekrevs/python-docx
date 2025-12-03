# Current State: What Works and What Doesn't

**Last updated:** 2024-12-03 (after T-TEST-01 baseline verification)

This document provides a quick reference for what python-docx can and cannot do.

## Executive Summary

python-docx 1.2.0 handles **basic document operations well** but cannot access **structured/dynamic content** like fields, content controls, footnotes, or bookmarks. These elements are preserved during round-trip but are invisible to the API.

**Coverage estimate:** ~60% for basic content, ~0% for structured content.

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

# Add new image
doc.add_picture('image.png', width=Inches(2))
```

### Comments
```python
for comment in doc.comments:
    comment.author            # ✅ Author name
    comment.text              # ✅ Comment text
    comment.date              # ✅ Date/time

# Add new comment
doc.add_comment("Comment text", author="Name", start=run1, end=run2)
```

---

## What Doesn't Work (No API Access)

### Fields — ❌ NOT ACCESSIBLE
```python
# Fields in document: PAGE, NUMPAGES, TOC, REF, DATE, etc.
#
# Problem: If paragraph contains "Page {PAGE} of {NUMPAGES}"
# You get: p.text == "Page  of "  (field results invisible!)
#
# No API for:
doc.fields                    # ❌ Doesn't exist
p.add_field('PAGE')           # ❌ Doesn't exist
```

**Found in test docs:** 1299 complex fields in one document

### Content Controls (SDT) — ❌ NOT ACCESSIBLE
```python
# Content controls are form fields: text boxes, dropdowns, checkboxes, date pickers
#
# No API for:
doc.content_controls          # ❌ Doesn't exist
sdt.value                     # ❌ Can't read values
sdt.tag                       # ❌ Can't read tags
p.add_content_control()       # ❌ Doesn't exist
```

**Found in test docs:** 63 content controls in one document

### Bookmarks — ❌ NOT ACCESSIBLE
```python
# Bookmarks are named locations in the document
#
# No API for:
doc.bookmarks                 # ❌ Doesn't exist
doc.bookmarks['name']         # ❌ Can't navigate to bookmark
p.add_bookmark('name')        # ❌ Doesn't exist
```

**Found in test docs:** 348 bookmarks in one document

### Footnotes & Endnotes — ❌ NOT ACCESSIBLE
```python
# Footnotes appear as reference marks with content at page bottom
#
# No API for:
doc.footnotes                 # ❌ Doesn't exist
footnote.text                 # ❌ Can't read footnote content
p.add_footnote("text")        # ❌ Doesn't exist
```

**Found in test docs:** 14 footnotes in test documents

### Track Changes — ❌ NOT ACCESSIBLE
```python
# Track changes show insertions, deletions, formatting changes
#
# No API for:
doc.revisions                 # ❌ Doesn't exist
revision.accept()             # ❌ Can't accept changes
revision.reject()             # ❌ Can't reject changes
```

**Found in test docs:** 43 tracked changes (33 insertions, 10 deletions)

### Numbering Definitions — ⚠️ PARTIAL
```python
# Can apply existing numbering:
p.style = 'List Bullet'       # ✅ Works

# Cannot create new numbering definitions:
doc.numbering.add_definition()  # ❌ Raises NotImplementedError
```

### Other Missing Features
```python
# Floating shapes/text boxes
doc.shapes                    # ❌ Doesn't exist (only inline_shapes)

# Math equations
doc.equations                 # ❌ Doesn't exist

# Charts (beyond detection)
chart.data                    # ❌ Can't read/modify chart data

# SmartArt (beyond detection)
smartart.text                 # ❌ Can't read/modify SmartArt
```

---

## The Workaround: Raw XML Access

Everything IS in the document — python-docx just doesn't expose it. You can access raw XML:

```python
from lxml import etree

doc = Document('file.docx')
body = doc.element.body

# Convert to string and re-parse for proper XPath
xml_str = etree.tostring(body)
tree = etree.fromstring(xml_str)

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# Find content controls
sdts = tree.xpath('.//w:sdt', namespaces=ns)
print(f"Found {len(sdts)} content controls")

# Find fields
fields = tree.xpath('.//w:fldChar', namespaces=ns)
print(f"Found {len(fields)} field characters")

# Find bookmarks
bookmarks = tree.xpath('.//w:bookmarkStart', namespaces=ns)
print(f"Found {len(bookmarks)} bookmarks")
```

**Problems with this approach:**
- Cumbersome, requires XML knowledge
- No convenient modification methods
- Easy to create invalid documents
- No type safety or validation

---

## Preservation Behavior

The good news: unsupported elements **survive round-trip**:

```python
doc = Document('complex.docx')  # Has fields, SDT, bookmarks...
doc.add_paragraph('New text')   # Make a change
doc.save('output.docx')         # Save

# Result: output.docx still has all fields, SDT, bookmarks intact!
# They're preserved, just not accessible via API
```

---

## Priority for WOTAN Development

Based on frequency in real-world test documents:

| Priority | Feature | Count in Test Docs | Backlog Item |
|----------|---------|-------------------|--------------|
| 1 | Complex Fields | 1299 | B-FLD-01, B-FLD-02 |
| 2 | Bookmarks | 348 | B-DRW-01 |
| 3 | Content Controls | 63 | B-SDT-01 |
| 4 | Track Changes | 43 | B-REV-01 |
| 5 | Footnotes | 14 | B-FN-01 |

---

## Quick Reference Card

| Feature | Read | Modify | Create |
|---------|------|--------|--------|
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
| **Fields** | ❌ | ❌ | ❌ |
| **Content Controls** | ❌ | ❌ | ❌ |
| **Bookmarks** | ❌ | ❌ | ❌ |
| **Footnotes** | ❌ | ❌ | ❌ |
| **Track Changes** | ❌ | ❌ | ❌ |
| **Numbering (new)** | ✅ | ⚠️ | ❌ |
| **Floating Shapes** | ⚠️ | ❌ | ❌ |
| **Math/Charts/SmartArt** | 🔍 | ❌ | ❌ |

**Legend:** ✅ Full | ⚠️ Partial | 🔍 Detect only | ❌ None
