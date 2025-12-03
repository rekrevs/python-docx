# python-docx (xtend fork)

This is an **experimental fork** of [python-docx](https://github.com/python-openxml/python-docx) with extensions to handle more of the full OOXML (.docx) format.

The upstream *python-docx* library provides excellent support for basic document operations. This fork extends it with read support for advanced features commonly found in real-world Word documents.

## Extensions

This fork adds the following capabilities:

| Feature | Read | Write | Description |
|---------|:----:|:-----:|-------------|
| **Content Controls (SDT)** | ✓ | ✓ | Structured document tags - text, date, dropdown, comboBox types |
| **Fields** | ✓ | ✓ | Simple and complex fields - PAGE, DATE, TOC, REF, HYPERLINK, etc. |
| **Footnotes & Endnotes** | ✓ | ✓ | Full paragraph and table support in notes |
| **Bookmarks** | ✓ | ✓ | Named locations for cross-references |
| **Track Changes** | ✓ | ✓ | Insertions/deletions with accept/reject support |
| **Floating Images** | ✓ | ✓ | Anchored shapes with position, wrap, and z-order |
| **Text Boxes** | ✓ | | Content in mc:AlternateContent elements |
| **Themes** | ✓ | ✓ | Theme colors and fonts (read and modify) |
| **Comments** | ✓ | ✓ | Comment threads with author metadata |
| **SVG Images** | ✓ | | Recognition and parsing of SVG files |
| **Modern Namespaces** | ✓ | | Word 2013+ namespaces (w14, w15, w16, etc.) |
| **Conformance Detection** | ✓ | | Detect Strict vs Transitional, Word version |

## Installation

```bash
pip install git+https://github.com/sverker/python-docx.git@xtend
```

## Example

```python
>>> from docx import Document

>>> document = Document()
>>> document.add_paragraph("It was a dark and stormy night.")
<docx.text.paragraph.Paragraph object at 0x10f19e760>
>>> document.save("dark-and-stormy.docx")

>>> document = Document("dark-and-stormy.docx")
>>> document.paragraphs[0].text
'It was a dark and stormy night.'
```

### Extension Examples

```python
>>> from docx import Document
>>> doc = Document("complex-document.docx")

# Access content controls
>>> for cc in doc.content_controls:
...     print(f"{cc.tag}: {cc.text}")

# Access fields
>>> for field in doc.fields:
...     print(f"{field.field_type}: {field.field_code}")

# Access footnotes
>>> for fn in doc.footnotes:
...     print(fn.paragraphs[0].text)

# Access bookmarks
>>> for bm in doc.bookmarks:
...     print(f"{bm.name}: {bm.id}")

# Access track changes
>>> for rev in doc.revisions.insertions:
...     print(f"{rev.author}: {rev.text}")

# Access theme
>>> print(doc.theme.colors.accent1)  # RGBColor
>>> print(doc.theme.fonts.minor_latin)  # Font name
```

## Documentation

For core python-docx functionality, see the [python-docx documentation](https://python-docx.readthedocs.org/en/latest/).

For a comprehensive API reference including all extensions, see **[WOTAN/docs/python-docx-api.md](WOTAN/docs/python-docx-api.md)**. This includes:
- Complete API reference for all features
- Pragmatics section with real-world document patterns
- Tips for handling complex documents (nested content controls, machine-generated files, etc.)

## Status

This is an experimental fork. All original python-docx tests pass. Extensions are additive and should not break existing functionality.
