# python-docx (xtend fork)

This is an **experimental fork** of [python-docx](https://github.com/python-openxml/python-docx) with extensions to handle more of the full OOXML (.docx) format.

The upstream *python-docx* library provides excellent support for basic document operations. This fork extends it with read support for advanced features commonly found in real-world Word documents.

## Extensions

This fork adds the following capabilities:

| Feature | Description |
|---------|-------------|
| **Content Controls (SDT)** | Read structured document tags - text, date, dropdown, comboBox types. Access by tag/alias. |
| **Fields** | Read simple and complex fields - PAGE, DATE, TOC, REF, CITATION, HYPERLINK, etc. |
| **Footnotes & Endnotes** | Read footnote/endnote content with full paragraph and table support. |
| **Bookmarks** | Read bookmarks, filter by type (TOC, Ref, user-defined). |
| **Track Changes** | Read-only access to insertions and deletions with author/date metadata. |
| **Floating Images** | Read anchored (floating) shapes with position and size properties. |
| **Themes** | Read theme colors and fonts from the document theme. |
| **SVG Images** | Recognition and parsing of SVG image files. |
| **Modern Namespaces** | Support for Word 2013+ namespaces (w14, w15, w16, etc.). |

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

Extension features are documented in the `WOTAN/` directory of this repository.

## Status

This is an experimental fork. The extensions focus on **read support** for advanced features. Write/create support for these features is not yet implemented.

All original python-docx tests pass. Extensions are additive and should not break existing functionality.
