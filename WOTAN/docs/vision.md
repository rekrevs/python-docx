# Vision: Complete python-docx

## The Goal

**Total document mastery**: Read any .docx file, parse it into a complete object model where every element is accessible and manipulable, transform it in any way desired, and write it back as a valid document.

```
.docx file → Complete Object Model → Arbitrary Manipulation → .docx file
```

This means:
- **No black boxes**: Every element in the document is exposed, not just "supported" elements
- **Full fidelity**: Nothing is silently dropped or ignored
- **Bidirectional**: Read and write with equal capability
- **Composable**: Build documents from scratch or tear them apart and reassemble

## Current Reality vs. Vision

### What Works Now

python-docx currently provides a **high-level API for common operations**:

| Capability | Read | Modify | Create |
|------------|------|--------|--------|
| Paragraphs & text | ✅ | ✅ | ✅ |
| Character formatting | ✅ | ✅ | ✅ |
| Paragraph formatting | ✅ | ✅ | ✅ |
| Tables | ✅ | ✅ | ✅ |
| Styles | ✅ | ✅ | ✅ |
| Headers/footers | ✅ | ✅ | ✅ |
| Sections | ✅ | ✅ | ✅ |
| Inline images | ✅ | ✅ | ✅ |
| Hyperlinks | ✅ | ✅ | ✅ |
| Comments | ✅ | ✅ | ✅ |
| Core properties | ✅ | ✅ | ✅ |

### What's Now Implemented (WOTAN Extensions)

These elements are now accessible through the extended API:

| Element | In Test Docs | Read | Modify | Create |
|---------|--------------|------|--------|--------|
| **Fields** (TOC, page#, cross-ref) | 1299 | ✅ | ⚠️ | ✅ |
| **Content Controls** (SDT) | 63 | ✅ | ✅ | ✅ |
| **Bookmarks** | 348 | ✅ | ❌ | ✅ |
| **Footnotes/Endnotes** | 14 | ✅ | ✅ | ✅ |
| **Track Changes** | 43 | ✅ | ✅ | ❌ |
| **Floating shapes** | 86+ | ✅ | ⚠️ | ✅ |
| **Text boxes** | 3+ | ✅ | ✅ | ❌ |
| **Theme colors/fonts** | all | ✅ | ✅ | ❌ |
| **SVG images** | ? | ✅ | ❌ | ❌ |
| **Conformance detection** | n/a | ✅ | n/a | n/a |

### Remaining Gaps

| Element | Read | Modify | Create |
|---------|------|--------|--------|
| **Math equations** | ❌ | ❌ | ❌ |
| **Charts** | 🔍 | ❌ | ❌ |
| **SmartArt** | 🔍 | ❌ | ❌ |
| **Custom XML** | ❌ | ❌ | ❌ |
| **Text box creation** | n/a | n/a | ❌ |
| **Bookmark modification** | n/a | ❌ | n/a |

**Legend**: ✅ Full | ⚠️ Partial | 🔍 Detect only | ❌ None

### The Preservation Behavior

python-docx **preserves** all elements during round-trip:
```python
doc = Document('complex.docx')  # Has fields, SDT, bookmarks...
doc.add_paragraph('New text')
doc.save('output.docx')  # Everything still there!
```

With WOTAN extensions, you can now:
- ✅ Read what a field contains (`doc.fields`)
- ✅ Modify a content control's value (`doc.content_controls`)
- ✅ Navigate to a bookmark (`doc.bookmarks`)
- ✅ Accept/reject a tracked change (`doc.revisions`)
- ✅ Extract footnote text (`doc.footnotes`)

## What "Complete" Means

### Level 1: Full Read Access
Every element in the document is accessible through the API:
```python
doc = Document('complex.docx')

# Fields
for field in doc.fields:
    print(f"{field.type}: {field.code} = {field.result}")

# Content controls
for sdt in doc.content_controls:
    print(f"{sdt.tag}: {sdt.value}")

# Bookmarks
for bookmark in doc.bookmarks:
    print(f"{bookmark.name} at {bookmark.start}")

# Footnotes
for footnote in doc.footnotes:
    print(f"[{footnote.id}] {footnote.text}")

# Track changes
for revision in doc.revisions:
    print(f"{revision.type} by {revision.author}: {revision.text}")
```

### Level 2: Full Modification
Every element can be modified:
```python
# Update a content control
doc.content_controls['customer_name'].value = "ACME Corp"

# Accept all changes from a specific author
for rev in doc.revisions:
    if rev.author == "Bob":
        rev.accept()

# Update field results
doc.update_fields()

# Move a bookmark
doc.bookmarks['section_start'].move_to(doc.paragraphs[10])
```

### Level 3: Full Creation
Every element can be created from scratch:
```python
doc = Document()

# Add a field
doc.add_paragraph().add_field('PAGE')

# Add a content control
p = doc.add_paragraph()
sdt = p.add_content_control(type='text', tag='user_input')
sdt.placeholder = "Enter name..."

# Add a footnote
p = doc.add_paragraph("See note")
p.add_footnote("This is the footnote text.")

# Add a bookmark
doc.add_bookmark('important_section', start=para1, end=para5)

# Create numbered list from scratch
num_def = doc.numbering.add_definition(style='decimal')
doc.add_paragraph("First item", style='ListNumber')
```

### Level 4: Low-Level Access
Direct access to underlying XML when needed:
```python
# Get raw XML for any element
xml = doc.paragraphs[0].element.xml

# Insert raw XML
from docx.oxml import parse_xml
custom_element = parse_xml('<w:customXml>...</w:customXml>')
doc.element.body.append(custom_element)

# Query with XPath
sdts = doc.element.body.xpath('.//w:sdt')
```

## Architecture Principle

The key insight: **two APIs, one model**

1. **High-level API** (current python-docx style)
   - Convenient, Pythonic
   - Hides XML complexity
   - Good for common operations

2. **Low-level API** (element access)
   - Full XML access
   - Nothing hidden
   - Good for advanced/unusual operations

Both operate on the same underlying object model. The high-level API is built on top of the low-level access, not instead of it.

## Success Criteria

The project achieves its vision when:

1. **Parse completely**: `Document('any.docx')` exposes 100% of document content
2. **Modify anything**: Any exposed element can be changed
3. **Create anything**: Any valid OOXML structure can be created programmatically
4. **Round-trip perfectly**: Open → Save produces identical output (when no changes made)
5. **Fail explicitly**: Unknown elements raise warnings/errors, never silent data loss

## Non-Goals

- **Rendering**: We don't display documents, we manipulate structure
- **Binary .doc**: Only .docx (OOXML) format
- **Other Office formats**: Focus on Word, not Excel/PowerPoint
- **Word behavior emulation**: We follow the spec, not Word's quirks

## Approach

1. **Inventory first**: Map every OOXML element to its python-docx status
2. **Read before write**: Implement read access before create/modify
3. **Test with real docs**: Use `WOTAN/example-docs/` as acceptance criteria
4. **Layered implementation**: OXML classes → Proxy objects → High-level API
5. **Preserve compatibility**: Don't break existing code
