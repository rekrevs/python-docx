# python-docx Architecture

This document describes the existing architecture of python-docx and how extensions should integrate.

## Layer Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Public API Layer                         │
│  Document(), document.add_paragraph(), paragraph.text, etc. │
│  Location: api.py, document.py                              │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                  Proxy Object Layer                         │
│  Document, Paragraph, Run, Table, Section, Style, Font      │
│  Location: document.py, text/, table.py, section.py, etc.   │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                     OXML Layer                              │
│  CT_Document, CT_P, CT_R, CT_Tbl, CT_RPr, etc.             │
│  Declarative XML element classes using xmlchemy             │
│  Location: oxml/                                            │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                    lxml Layer                               │
│  BaseOxmlElement extends lxml.etree.ElementBase             │
│  Custom element class lookup                                │
│  Location: oxml/xmlchemy.py, oxml/parser.py                 │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                     OPC Layer                               │
│  Package, Part, Relationship, ContentType                   │
│  ZIP file handling, part serialization                      │
│  Location: opc/                                             │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│                   .docx File (ZIP)                          │
│  [Content_Types].xml, _rels/.rels, word/document.xml, etc.  │
└─────────────────────────────────────────────────────────────┘
```

## Key Components

### 1. Public API (`api.py`)

Entry point for users:

```python
from docx import Document

doc = Document()  # New document
doc = Document('existing.docx')  # Open existing
doc.save('output.docx')
```

The `Document()` function creates a `Document` proxy object backed by a `DocumentPart`.

### 2. Proxy Objects

User-facing classes that wrap OXML elements, providing Pythonic APIs:

| Proxy Class | OXML Class | Element |
|------------|------------|---------|
| `Document` | `CT_Document` | `w:document` |
| `Paragraph` | `CT_P` | `w:p` |
| `Run` | `CT_R` | `w:r` |
| `Table` | `CT_Tbl` | `w:tbl` |
| `Font` | `CT_RPr` | `w:rPr` |
| `ParagraphFormat` | `CT_PPr` | `w:pPr` |
| `Section` | `CT_SectPr` | `w:sectPr` |
| `Style` | `CT_Style` | `w:style` |

Proxy objects:
- Hold a reference to their OXML element (`self._element` or similar)
- Provide property accessors that read/write XML attributes
- Handle None/missing element cases gracefully
- Often inherit from `Parented` to access parent part

### 3. OXML Layer (`oxml/`)

Custom lxml element classes using declarative descriptors:

```python
class CT_P(BaseOxmlElement):
    """<w:p> element - paragraph"""

    pPr: CT_PPr | None = ZeroOrOne("w:pPr", successors=_tag_seq[1:])
    r_lst: list[CT_R] = ZeroOrMore("w:r", successors=_tag_seq[5:])

    def add_r(self) -> CT_R:
        """Add a new <w:r> child element."""
        r = OxmlElement("w:r")
        self.append(r)
        return r
```

Key patterns:
- `ZeroOrOne` — Optional single child element
- `ZeroOrMore` — List of child elements
- `OneAndOnlyOne` — Required single child
- `RequiredAttribute` / `OptionalAttribute` — XML attributes
- `successors` — Maintains correct element ordering per schema

### 4. xmlchemy (`oxml/xmlchemy.py`)

The declarative descriptor system (25K+ lines):

- `BaseOxmlElement` — Base class for all custom elements
- Descriptors auto-generate getters/setters
- Handles element creation, insertion at correct position
- `_add_*`, `_remove_*`, `get_or_add_*` methods generated

### 5. OPC Layer (`opc/`)

Open Packaging Conventions implementation:

| Component | Purpose |
|-----------|---------|
| `OpcPackage` | Top-level package (ZIP file) |
| `Part` / `XmlPart` | Individual parts within package |
| `Relationship` | Links between parts |
| `PackURI` | Part addressing |
| `pkgreader.py` | Unmarshalling from ZIP |
| `pkgwriter.py` | Serializing to ZIP |

### 6. Parts (`parts/`)

Concrete part implementations:

| Part Class | Content Type | File |
|------------|--------------|------|
| `DocumentPart` | `wordprocessingml.document.main` | word/document.xml |
| `StylesPart` | `wordprocessingml.styles` | word/styles.xml |
| `NumberingPart` | `wordprocessingml.numbering` | word/numbering.xml |
| `SettingsPart` | `wordprocessingml.settings` | word/settings.xml |
| `HeaderPart` | `wordprocessingml.header` | word/header*.xml |
| `FooterPart` | `wordprocessingml.footer` | word/footer*.xml |
| `CommentsPart` | `wordprocessingml.comments` | word/comments.xml |
| `ImagePart` | `image/*` | word/media/* |

Parts inherit from `XmlPart` and provide:
- `_element` — Root OXML element
- Relationship management
- Lazy loading of related parts

---

## Adding New Features

### Pattern: New Element Type

1. **Define OXML class** in `oxml/`:

```python
# oxml/sdt.py
class CT_SdtPr(BaseOxmlElement):
    """<w:sdtPr> element - content control properties"""

    tag: CT_String | None = ZeroOrOne("w:tag")
    alias: CT_String | None = ZeroOrOne("w:alias")
    # ... more properties
```

2. **Register element** in `oxml/__init__.py`:

```python
from .sdt import CT_SdtPr, CT_SdtBlock, CT_SdtContent

register_element_cls("w:sdtPr", CT_SdtPr)
register_element_cls("w:sdt", CT_SdtBlock)
register_element_cls("w:sdtContent", CT_SdtContent)
```

3. **Create proxy class**:

```python
# sdt.py
class ContentControl:
    """Proxy for a content control (SDT)."""

    def __init__(self, sdt_elm: CT_SdtBlock):
        self._element = sdt_elm

    @property
    def tag(self) -> str | None:
        sdtPr = self._element.sdtPr
        return sdtPr.tag.val if sdtPr and sdtPr.tag else None
```

4. **Integrate with existing API**:

```python
# In Paragraph or Document
@property
def content_controls(self) -> list[ContentControl]:
    return [ContentControl(sdt) for sdt in self._element.sdt_lst]
```

5. **Add tests**:
   - Unit tests in `tests/`
   - Acceptance tests in `features/`

### Pattern: New Part Type

1. **Define part class** in `parts/`:

```python
# parts/footnotes.py
class FootnotesPart(XmlPart):
    """Proxy for footnotes.xml part."""

    @classmethod
    def new(cls, package: OpcPackage) -> FootnotesPart:
        # Create from template or minimal XML
        ...

    @property
    def footnotes(self) -> list[Footnote]:
        return [Footnote(fn) for fn in self._element.footnote_lst]
```

2. **Add content type** to `opc/constants.py` if not present

3. **Wire into DocumentPart**:

```python
# parts/document.py
@property
def footnotes_part(self) -> FootnotesPart:
    try:
        return self.part_related_by(RT.FOOTNOTES)
    except KeyError:
        # Create if not exists
        ...
```

### Pattern: New Namespace

1. **Add to namespace map** in `oxml/ns.py`:

```python
nsmap = {
    # ... existing
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
}
```

2. **Use in OXML classes**:

```python
class CT_Color(BaseOxmlElement):
    # w14 theme color
    themeColor = OptionalAttribute("w14:themeColor", ST_String)
```

---

## File Layout

```
src/docx/
├── __init__.py          # Package init, version
├── api.py               # Document() entry point
├── document.py          # Document proxy class
├── blkcntnr.py          # BlockItemContainer base
├── section.py           # Section, Header, Footer
├── table.py             # Table, Row, Cell
├── shape.py             # InlineShape
├── drawing.py           # Drawing container
├── comments.py          # Comments
├── settings.py          # Settings proxy
├── shared.py            # Length, Parented, utilities
├── exceptions.py        # Custom exceptions
├── types.py             # Type definitions
│
├── text/                # Text-related proxies
│   ├── paragraph.py     # Paragraph
│   ├── run.py           # Run
│   ├── font.py          # Font (character formatting)
│   ├── parfmt.py        # ParagraphFormat
│   └── hyperlink.py     # Hyperlink
│
├── styles/              # Style system
│   ├── styles.py        # Styles collection
│   ├── style.py         # Style classes
│   └── latent.py        # Latent styles
│
├── image/               # Image handling
│   ├── image.py         # Image base
│   ├── png.py           # PNG parser
│   ├── jpeg.py          # JPEG parser
│   └── ...
│
├── oxml/                # XML element layer
│   ├── __init__.py      # Element registration
│   ├── xmlchemy.py      # Declarative descriptors
│   ├── parser.py        # XML parsing
│   ├── ns.py            # Namespaces
│   ├── simpletypes.py   # Simple type validators
│   ├── document.py      # CT_Document, CT_Body
│   ├── table.py         # CT_Tbl, CT_Row, CT_Tc
│   ├── section.py       # CT_SectPr
│   ├── styles.py        # CT_Styles, CT_Style
│   ├── numbering.py     # CT_Numbering
│   ├── settings.py      # CT_Settings
│   ├── comments.py      # CT_Comments
│   ├── drawing.py       # CT_Drawing
│   ├── shape.py         # CT_Inline, CT_Anchor
│   └── text/            # Text elements
│       ├── paragraph.py # CT_P
│       ├── run.py       # CT_R
│       ├── font.py      # CT_RPr
│       └── parfmt.py    # CT_PPr
│
├── opc/                 # OPC layer
│   ├── package.py       # OpcPackage
│   ├── part.py          # Part, XmlPart
│   ├── rel.py           # Relationships
│   ├── packuri.py       # Pack URIs
│   ├── pkgreader.py     # Read from ZIP
│   ├── pkgwriter.py     # Write to ZIP
│   └── constants.py     # Content types, rel types
│
├── parts/               # Document parts
│   ├── document.py      # DocumentPart
│   ├── styles.py        # StylesPart
│   ├── numbering.py     # NumberingPart
│   ├── settings.py      # SettingsPart
│   ├── hdrftr.py        # HeaderPart, FooterPart
│   ├── comments.py      # CommentsPart
│   ├── image.py         # ImagePart
│   └── story.py         # StoryPart base
│
└── templates/           # Default document templates
    └── default-docx-template/
```

---

## Design Principles

1. **Proxy pattern**: User-facing objects wrap XML elements
2. **Lazy evaluation**: Don't create XML until needed
3. **None propagation**: Missing elements return None, not errors
4. **Immutable XML ordering**: Use `successors` to maintain schema order
5. **Part isolation**: Each part manages its own relationships
6. **Template-based creation**: New documents start from templates
