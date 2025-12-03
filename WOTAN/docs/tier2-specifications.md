# Tier 2 Implementation Specifications: Write Support for Extensions

**Created:** 2025-12-03
**Status:** READY FOR IMPLEMENTATION

This document provides specifications for all Tier 2 backlog items.

---

## B-SDT-02: Content Control Creation

### Objective

Enable creating new content controls (structured document tags) programmatically.

### OOXML Specification

Content controls use the `w:sdt` element with:
- `w:sdtPr` - Properties (alias, tag, type, lock settings)
- `w:sdtContent` - The actual content

#### SDT Types

| Type | Property Element | Description |
|------|------------------|-------------|
| Rich Text | (none/default) | Allows any content |
| Plain Text | `w:text` | Plain text only |
| Picture | `w:picture` | Single image |
| Combo Box | `w:comboBox` | Editable dropdown |
| Drop-Down List | `w:dropDownList` | Fixed list selection |
| Date Picker | `w:date` | Date selection |
| Checkbox | `w14:checkbox` | Checkbox (Word 2010+) |
| Building Block | `w:docPartObj` | Document building block |

#### SDT Levels

- **Block-level**: Contains paragraphs, tables
- **Inline-level**: Contains runs within a paragraph
- **Row-level**: Contains a table row
- **Cell-level**: Contains a table cell

### Current State

- Can read content controls via `doc.content_controls`
- Can access content via `SdtBlockContentControl.paragraphs`, `.tables`
- Cannot create new content controls

### Proposed API

```python
# Create a block-level text content control
sdt = document.add_content_control(
    sdt_type='text',  # or 'rich_text', 'date', 'dropdown', 'combobox', 'checkbox'
    tag='customer_name',
    alias='Customer Name'
)
sdt.add_paragraph("Default placeholder text")

# Create a dropdown content control
sdt = document.add_content_control(
    sdt_type='dropdown',
    tag='status',
    alias='Status'
)
sdt.add_list_item('Pending', value='pending')
sdt.add_list_item('Approved', value='approved')
sdt.add_list_item('Rejected', value='rejected')

# Create a date content control
sdt = document.add_content_control(
    sdt_type='date',
    tag='due_date',
    alias='Due Date',
    date_format='yyyy-MM-dd'
)

# Add content control inline (within a paragraph)
para = document.add_paragraph("Customer: ")
sdt = para.add_content_control(sdt_type='text', tag='name')
sdt.add_run("Enter name")
```

### Implementation Plan

1. Add `add_content_control()` method to `Document` class
2. Add `add_content_control()` method to `Paragraph` class (for inline SDT)
3. Create OXML helper methods in `CT_SdtPr` for setting properties
4. Create `CT_SdtBlock.new()` and `CT_SdtRun.new()` factory methods
5. Handle dropdown/combobox list items via `w:listItem` elements
6. Handle date properties via `w:date` element with format attributes

### XML Examples

**Block-level text SDT:**
```xml
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="Customer Name"/>
    <w:tag w:val="customer_name"/>
    <w:text/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r><w:t>Default text</w:t></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
```

**Dropdown SDT:**
```xml
<w:sdt>
  <w:sdtPr>
    <w:alias w:val="Status"/>
    <w:tag w:val="status"/>
    <w:dropDownList>
      <w:listItem w:displayText="Pending" w:value="pending"/>
      <w:listItem w:displayText="Approved" w:value="approved"/>
    </w:dropDownList>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:r><w:t>Pending</w:t></w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>
```

### Acceptance Criteria

- [ ] Create block-level content controls (text, rich text, date, dropdown, combobox)
- [ ] Create inline content controls within paragraphs
- [ ] Set tag, alias, and type-specific properties
- [ ] Add list items to dropdown/combobox controls
- [ ] Set date format for date controls
- [ ] Document opens correctly in Word with functional controls

### References

- [SDT Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Structured_topic_ID0ERWJS.html)
- [sdtPr Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_sdtPr_topic_ID0EOSZS.html)

---

## B-FLD-03: Field Creation

### Objective

Enable creating new fields (PAGE, DATE, REF, etc.) programmatically.

### OOXML Specification

Fields can be represented in two ways:

#### Simple Fields (`w:fldSimple`)

Single element containing field instruction and result:
```xml
<w:fldSimple w:instr=" PAGE ">
  <w:r><w:t>1</w:t></w:r>
</w:fldSimple>
```

#### Complex Fields (`w:fldChar` + `w:instrText`)

Multi-element structure for complex formatting:
```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

### Common Field Types

| Field | Purpose | Example Instruction |
|-------|---------|---------------------|
| PAGE | Current page number | `PAGE` |
| NUMPAGES | Total pages | `NUMPAGES` |
| DATE | Current date | `DATE \@ "yyyy-MM-dd"` |
| TIME | Current time | `TIME \@ "HH:mm"` |
| REF | Cross-reference | `REF _Ref123 \h` |
| PAGEREF | Page of bookmark | `PAGEREF _Toc123 \h` |
| HYPERLINK | Hyperlink | `HYPERLINK "url"` |
| TOC | Table of contents | `TOC \o "1-3"` |

### Current State

- Can read simple and complex fields via `doc.fields`
- Can filter by type, access field code and result
- Cannot create new fields

### Proposed API

```python
# Add simple field to a run
run = paragraph.add_run()
run.add_field('PAGE')  # Creates simple field

# Add field with switches
run.add_field('DATE', switches=r'\@ "MMMM d, yyyy"')

# Add complex field (when formatting differs)
run.add_complex_field('REF', '_Ref123456', switches=r'\h')

# Convenience methods on paragraph
paragraph.add_page_number()
paragraph.add_page_count()
paragraph.add_date(format='yyyy-MM-dd')
```

### Implementation Plan

1. Add `add_field()` method to `Run` class
2. Create `CT_FldSimple.new()` factory method
3. Create helper for complex field structure
4. Add convenience methods to `Paragraph` class
5. Handle field switches/formatting

### Acceptance Criteria

- [ ] Create simple fields (PAGE, NUMPAGES, DATE, TIME)
- [ ] Create fields with format switches
- [ ] Create cross-reference fields (REF, PAGEREF)
- [ ] Fields display placeholder until updated in Word
- [ ] Fields update correctly when opened in Word

### References

- [Field XML Representation](https://ooxml.info/docs/17/17.16/17.16.2/)
- [Field Instructions](http://www.officeopenxml.com/WPfieldInstructions.php)
- [fldSimple Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_fldSimple_topic_ID0EPAV1.html)

---

## B-DRW-04: Bookmark Creation

### Objective

Enable creating new bookmarks programmatically.

### OOXML Specification

Bookmarks use paired empty elements:
- `w:bookmarkStart` - Start marker with `w:id` and `w:name`
- `w:bookmarkEnd` - End marker with matching `w:id`

```xml
<w:bookmarkStart w:id="0" w:name="MyBookmark"/>
<w:p>
  <w:r><w:t>Bookmarked content</w:t></w:r>
</w:p>
<w:bookmarkEnd w:id="0"/>
```

#### ID Requirements

- IDs must be unique among all `bookmarkStart` elements in the document
- Each `bookmarkStart` must have a corresponding `bookmarkEnd` with matching ID
- IDs are integers (typically 0-based sequential)

#### Point Bookmarks

Start and end at same location (no content spanned):
```xml
<w:bookmarkStart w:id="0" w:name="InsertPoint"/>
<w:bookmarkEnd w:id="0"/>
```

### Current State

- Can read bookmarks via `doc.bookmarks`
- Can access by name, filter by type
- Cannot create new bookmarks

### Proposed API

```python
# Create a bookmark around a paragraph
para = document.paragraphs[0]
bookmark = document.add_bookmark('chapter1', start=para, end=para)

# Create a bookmark around a range
start_para = document.paragraphs[0]
end_para = document.paragraphs[2]
bookmark = document.add_bookmark('section1', start=start_para, end=end_para)

# Create a point bookmark (no range)
bookmark = document.add_bookmark('insert_point', at=para)

# Create bookmark at current position in a paragraph
run = para.add_run()
run.add_bookmark('ref_target')
```

### Implementation Plan

1. Add bookmark ID management to document (track max ID, allocate new)
2. Create `CT_Bookmark.new()` and `CT_MarkupRange.new()` factory methods
3. Add `add_bookmark()` method to `Document` class
4. Handle start/end element placement in document body
5. Validate unique bookmark names

### Acceptance Criteria

- [ ] Create named bookmarks around paragraphs
- [ ] Create bookmarks spanning multiple paragraphs
- [ ] Create point bookmarks
- [ ] Automatic unique ID generation
- [ ] Validate bookmark name uniqueness
- [ ] Bookmarks visible in Word Navigation pane

### References

- [Bookmarks Specification](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Bookmarks_topic_ID0EFMWW.html)
- [bookmarkStart](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_bookmarkStart_topic_ID0EHVXW.html)
- [OpenXML Bookmark ID Management](https://stackoverflow.com/questions/29772833/defining-correct-ids-for-new-bookmarks-in-openxml-document)

---

## B-FN-02: Footnote/Endnote Creation

### Objective

Enable creating new footnotes and endnotes programmatically.

### OOXML Specification

Footnotes involve two parts:

1. **Reference in document** (`w:footnoteReference` in a run):
```xml
<w:r>
  <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
  <w:footnoteReference w:id="2"/>
</w:r>
```

2. **Content in footnotes.xml** (`w:footnote`):
```xml
<w:footnote w:id="2">
  <w:p>
    <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r>
      <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
      <w:footnoteRef/>
    </w:r>
    <w:r>
      <w:t> This is the footnote text.</w:t>
    </w:r>
  </w:p>
</w:footnote>
```

#### Special Footnotes

- ID 0: Separator (horizontal line)
- ID 1: Continuation separator
- ID 2+: Regular footnotes

#### Automatic Numbering

The `w:footnoteRef` element in the footnote content creates the automatic numbering marker.

### Current State

- Can read footnotes/endnotes via `doc.footnotes`, `doc.endnotes`
- Can access content (paragraphs, tables)
- Cannot create new footnotes/endnotes

### Proposed API

```python
# Add footnote to a run
run = paragraph.add_run("Some text")
footnote = run.add_footnote("This is the footnote text.")

# Add footnote with multiple paragraphs
run = paragraph.add_run("Another point")
footnote = run.add_footnote()
footnote.add_paragraph("First paragraph of footnote.")
footnote.add_paragraph("Second paragraph of footnote.")

# Add endnote
run = paragraph.add_run("Citation needed")
endnote = run.add_endnote("Smith, John. Some Book. 2020.")
```

### Implementation Plan

1. Add footnote ID management to FootnotesPart
2. Create `add_footnote()` method on FootnotesPart
3. Add `add_footnote()` method to `Run` class
4. Create footnote reference run in document
5. Create footnote content in footnotes.xml
6. Handle endnotes similarly with EndnotesPart
7. Ensure correct styles applied (FootnoteText, FootnoteReference)

### Acceptance Criteria

- [ ] Create footnotes with text content
- [ ] Create footnotes with multiple paragraphs
- [ ] Create endnotes
- [ ] Automatic footnote numbering works in Word
- [ ] Correct styles applied to footnote text and reference

### References

- [Footnote Content](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_footnote_topic_ID0EKU5U.html)
- [Footnote Reference](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_footnoteReference_topic_ID0EZ4AV.html)
- [OpenXML SDK Footnote Creation](https://stackoverflow.com/questions/37112277/add-footnotes-to-word-document-programatically-using-openxml-sdk-c-sharp)

---

## B-REV-02: Accept/Reject Track Changes

### Objective

Programmatically accept or reject tracked changes (revisions).

### OOXML Specification

Track changes are represented by:

- `w:ins` - Inserted content
- `w:del` - Deleted content
- `w:rPrChange` - Run property changes
- `w:pPrChange` - Paragraph property changes
- `w:sectPrChange` - Section property changes
- `w:tblPrChange` - Table property changes

#### Accept Insertion

Remove `w:ins` wrapper, keep content:
```xml
<!-- Before -->
<w:ins w:author="John" w:date="2024-01-01T10:00:00Z">
  <w:r><w:t>New text</w:t></w:r>
</w:ins>

<!-- After accepting -->
<w:r><w:t>New text</w:t></w:r>
```

#### Reject Insertion

Remove `w:ins` and its content entirely.

#### Accept Deletion

Remove `w:del` and its content entirely.

#### Reject Deletion

Remove `w:del` wrapper, keep content:
```xml
<!-- Before -->
<w:del w:author="John" w:date="2024-01-01T10:00:00Z">
  <w:r><w:t>Old text</w:t></w:r>
</w:del>

<!-- After rejecting -->
<w:r><w:t>Old text</w:t></w:r>
```

### Current State

- Can read revisions via `doc.revisions`
- Can access insertions, deletions, author, date, text
- Cannot accept or reject changes

### Proposed API

```python
# Accept individual revision
for revision in doc.revisions:
    if revision.author == "John":
        revision.accept()

# Reject individual revision
revision = doc.revisions[0]
revision.reject()

# Accept all
doc.revisions.accept_all()

# Reject all
doc.revisions.reject_all()

# Accept all from specific author
doc.revisions.accept_all(author="John")

# Accept/reject by type
for revision in doc.revisions.insertions:
    revision.accept()
```

### Implementation Plan

1. Add `accept()` method to `Revision` class
2. Add `reject()` method to `Revision` class
3. Implement logic for each revision type:
   - `w:ins`: accept = unwrap, reject = remove
   - `w:del`: accept = remove, reject = unwrap
   - Property changes: accept = keep new, reject = restore old
4. Add `accept_all()` and `reject_all()` to `Revisions` collection
5. Handle nested revisions correctly

### Acceptance Criteria

- [ ] Accept individual insertions (text becomes normal)
- [ ] Reject individual insertions (text removed)
- [ ] Accept individual deletions (text removed)
- [ ] Reject individual deletions (text restored)
- [ ] Accept/reject property changes
- [ ] Bulk accept_all() / reject_all()
- [ ] Filter by author when accepting/rejecting

### References

- [Track Revisions](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_trackRevisions_topic_ID0EKXKY.html)
- [Eric White: Detecting Track Revisions](http://www.ericwhite.com/blog/using-xml-dom-to-detect-tracked-revisions-in-an-open-xml-wordprocessingml-document/)

---

## Implementation Order

Recommended order based on dependencies and complexity:

1. **B-FLD-03: Field Creation** - Standalone, simpler XML structure
2. **B-DRW-04: Bookmark Creation** - Standalone, needed for cross-references
3. **B-FN-02: Footnote/Endnote Creation** - Requires FootnotesPart coordination
4. **B-SDT-02: Content Control Creation** - More complex with multiple types
5. **B-REV-02: Accept/Reject Track Changes** - Most complex, many edge cases

---

## Testing Strategy

For each item:

1. **Unit tests** - Test OXML element creation
2. **Integration tests** - Test full API workflow
3. **Round-trip tests** - Create, save, reopen, verify
4. **Word compatibility tests** - Open in Word 2010/2016/2019/365, verify functionality
