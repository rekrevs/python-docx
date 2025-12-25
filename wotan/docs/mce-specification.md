# MCE (Markup Compatibility and Extensibility) Implementation Specification

**Document:** B-CORE-01 Implementation Specification
**Created:** 2025-12-03
**Status:** READY FOR IMPLEMENTATION

## Executive Summary

This specification defines how python-docx should handle `mc:AlternateContent` blocks to enable full round-trip document manipulation. Currently, content inside these blocks (text boxes, modern shapes) is invisible to the API.

## Problem Statement

### Current Behavior

Content inside `mc:AlternateContent` blocks is:
- **Preserved** in XML on save (round-trip works at XML level)
- **Invisible** to `doc.paragraphs`, `doc.tables`, and other API access
- **Not modifiable** through the standard API

### Affected Content

Based on analysis of real documents:
1. **Text boxes** - Very common, contain `w:txbxContent` with paragraphs
2. **Modern vector graphics** - DrawingML shapes (`wps:wsp`, `wpg:wgrp`)
3. **Word 2010+ features** - Content wrapped for backwards compatibility

### Evidence

From `Q-NEXUS_Application Form (Part B)_250925.docx`:
- 3 `mc:AlternateContent` elements containing text boxes
- Text "Instructions, please remove" is in a text box
- This text is **NOT** accessible via `doc.paragraphs` (confirmed)

## MCE Specification Overview

Based on ECMA-376 Part 3 / ISO 29500-3.

### Key Elements

| Element | Purpose |
|---------|---------|
| `mc:AlternateContent` | Container for version-specific alternatives |
| `mc:Choice` | Preferred content if `@Requires` namespaces are supported |
| `mc:Fallback` | Fallback content for older readers |

### Key Attributes

| Attribute | Purpose |
|-----------|---------|
| `mc:Ignorable` | Namespaces that can be ignored if unknown |
| `mc:MustUnderstand` | Namespaces that must be understood (halt if not) |
| `mc:ProcessContent` | Process children of ignored elements |

### Typical Structure (Text Box)

```xml
<mc:AlternateContent>
  <mc:Choice Requires="wps">
    <w:drawing>
      <wp:anchor ...>
        <a:graphic>
          <a:graphicData uri="...">
            <wps:wsp>
              <wps:txbx>
                <w:txbxContent>
                  <w:p><w:r><w:t>Text here</w:t></w:r></w:p>
                </w:txbxContent>
              </wps:txbx>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </mc:Choice>
  <mc:Fallback>
    <w:pict>
      <v:shape ...>
        <v:textbox>
          <w:txbxContent>
            <w:p><w:r><w:t>Text here</w:t></w:r></w:p>
          </w:txbxContent>
        </v:textbox>
      </v:shape>
    </w:pict>
  </mc:Fallback>
</mc:AlternateContent>
```

## How Other Libraries Handle MCE

### docx4j (Java)

- Uses preprocessing XSLT when JAXB encounters unexpected content
- **Historically selected mc:Fallback** (until v3.3.8)
- Now preserves mc:AlternateContent in runs
- **Warning**: If docx4j saves the document, Word 2010+ content is lost (effectively Word 2007 docx)

### Open XML SDK (.NET)

- Has explicit `AlternateContent` class
- Provides `GetContentFromACBlock(block, FileFormatVersions)` method
- Can target specific Office versions

### Community Workarounds (python-docx)

1. **XPath access**: `body.xpath('.//w:txbxContent//w:p')`
2. **Custom TextBox class**: Extend `BlockItemContainer` for `w:txbxContent`
3. **Modify run text**: Avoid clearing paragraphs, modify run text directly

## Implementation Strategy

### Recommended Approach: "Choice-First with Preservation"

1. **Reading**: Access content from `mc:Choice` (modern format)
2. **Writing**: Modify content in both `mc:Choice` and `mc:Fallback`
3. **Round-trip**: Preserve entire `mc:AlternateContent` structure

### Why Choice-First?

- `mc:Choice` contains the modern, richer representation
- `mc:Fallback` is for legacy compatibility
- Most python-docx users have Word 2010+

### Implementation Plan

#### Phase 1: Add `mc` namespace and MCE-aware element discovery

1. Add `mc` namespace to `nsmap` in `src/docx/oxml/ns.py`:
   ```python
   "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
   ```

2. Add VML namespace for fallback parsing:
   ```python
   "v": "urn:schemas-microsoft-com:vml",
   ```

#### Phase 2: Create MCE OXML layer

Create `src/docx/oxml/mce.py`:

```python
class CT_AlternateContent(BaseOxmlElement):
    """<mc:AlternateContent> element."""

    @property
    def choice(self) -> CT_Choice | None:
        """First mc:Choice child element."""

    @property
    def fallback(self) -> CT_Fallback | None:
        """mc:Fallback child element."""

    @property
    def requires(self) -> str | None:
        """Requires attribute from first Choice."""

class CT_Choice(BaseOxmlElement):
    """<mc:Choice> element."""

    @property
    def requires(self) -> str:
        """Required namespace prefix(es)."""

class CT_Fallback(BaseOxmlElement):
    """<mc:Fallback> element."""
```

#### Phase 3: Create TextBox proxy classes

Create `src/docx/textbox.py`:

```python
class TextBoxes:
    """Collection of text boxes in the document."""

    def __iter__(self) -> Iterator[TextBox]:
        """Iterate over all text boxes."""

    def __len__(self) -> int:
        """Number of text boxes."""

class TextBox(BlockItemContainer):
    """Proxy for a text box.

    A text box contains paragraphs and tables, accessed via
    the standard BlockItemContainer interface.
    """

    @property
    def paragraphs(self) -> list[Paragraph]:
        """Paragraphs in this text box."""

    @property
    def tables(self) -> list[Table]:
        """Tables in this text box."""

    # Position/size properties (from anchor)
    @property
    def width(self) -> Length | None: ...

    @property
    def height(self) -> Length | None: ...
```

#### Phase 4: Expose on Document

Add to `Document` class:

```python
@property
def text_boxes(self) -> TextBoxes:
    """Collection of text boxes in this document."""
    return TextBoxes(self._part)
```

#### Phase 5: Ensure dual-write for modifications

When modifying content in a text box:
1. Modify `w:txbxContent` in `mc:Choice`
2. Sync changes to `w:txbxContent` in `mc:Fallback`

This ensures the document remains compatible with older Word versions.

### Alternative Considerations

#### Option A: Flatten MCE (NOT RECOMMENDED)

Select content from `mc:Choice` and discard `mc:Fallback`.
- **Pro**: Simpler implementation
- **Con**: Breaks backwards compatibility, document may not open in Word 2007

#### Option B: Full MCE Processor (DEFERRED)

Implement complete ECMA-376 Part 3 MCE processing.
- **Pro**: Fully specification-compliant
- **Con**: Complex, probably unnecessary for most use cases

## API Design

### Reading Text Boxes

```python
from docx import Document

doc = Document("document.docx")

# Iterate over all text boxes
for textbox in doc.text_boxes:
    print(f"Text box with {len(textbox.paragraphs)} paragraphs")
    for para in textbox.paragraphs:
        print(f"  {para.text}")

# Access by index
first_box = doc.text_boxes[0]
```

### Modifying Text Box Content

```python
# Modify paragraph text
textbox = doc.text_boxes[0]
textbox.paragraphs[0].text = "New text"

# Add paragraph
textbox.add_paragraph("Additional content")

# Save - both Choice and Fallback are updated
doc.save("modified.docx")
```

### Future: Creating Text Boxes

```python
# Future enhancement (B-DRW-03)
textbox = doc.add_text_box(
    width=Inches(2),
    height=Inches(1),
    left=Inches(1),
    top=Inches(1)
)
textbox.add_paragraph("New text box content")
```

## Testing Strategy

### Unit Tests

1. Test MCE element parsing
2. Test text box discovery in documents with/without text boxes
3. Test content access through proxy
4. Test modification and dual-write

### Integration Tests

1. Open document with text boxes, verify content accessible
2. Modify text box content, save, reopen, verify changes
3. Open modified document in Word 2007/2010/2016, verify no corruption
4. Round-trip test: open, don't modify, save, verify identical

### Test Documents

- `WOTAN/example-docs/Q-NEXUS_Application Form (Part B)_250925.docx` - Has 3 text boxes
- Create minimal test document with single text box

## Acceptance Criteria

From B-CORE-01:

- [x] Document MCE specification ✓
- [ ] Paragraphs/tables inside mc:AlternateContent are accessible via existing APIs
- [ ] Text box content is readable and modifiable
- [ ] Round-trip preserves mc:AlternateContent structure
- [ ] Document can still be opened in older Word versions (mc:Fallback preserved)

## References

### Specifications
- [ECMA-376 Part 3](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Markup Compatibility and Extensibility
- [ISO 29500-3](https://www.loc.gov/preservation/digital/formats/fdd/fdd000396.shtml)

### Articles
- [Eric White: Markup Compatibility and Extensibility](http://www.ericwhite.com/blog/markup-compatibility-and-extensibility/)
- [Microsoft Learn: Introduction to markup compatibility](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility)
- [How to edit Microsoft Word documents in Python](https://www.rikvoorhaar.com/blog/python_docx) - TextBox workaround

### Issues
- [python-docx #1389](https://github.com/python-openxml/python-docx/issues/1389) - mc:AlternateContent content invisible
- [docx4j #443](https://github.com/plutext/docx4j/issues/443) - AlternateContent within RPr
