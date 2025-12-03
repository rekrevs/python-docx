# python-docx Extension Backlog

This backlog tracks planned enhancements to bring python-docx closer to full OOXML specification compliance.

## Categories

| Category | Description |
|----------|-------------|
| `CORE` | Core document structure (body, sections, paragraphs) |
| `FMT` | Formatting (character, paragraph, table formatting) |
| `TBL` | Tables |
| `IMG` | Images and shapes |
| `STY` | Styles |
| `OPC` | Open Packaging Conventions layer |
| `DOC` | Documentation |
| `TEST` | Test infrastructure |
| `SDT` | Structured Document Tags (Content Controls) |
| `FLD` | Fields |
| `NUM` | Numbering/Lists |
| `FN` | Footnotes/Endnotes |
| `REV` | Revisions/Track Changes |
| `DRW` | Drawing/Shapes |
| `NS` | Namespace extensions (w14, w15, w16, etc.) |

---

## Tier 0: Foundation

### B-TEST-02: Research Community Knowledge and Version Compatibility `[DONE]`

**Intent:** Research developer forums, known issues, and OOXML version compatibility before starting implementation.

**Next:** T-TEST-02 (completed)

**Details:**
- Search developer forums (Stack Overflow, GitHub issues, python-docx discussions)
- Understand OOXML versioning (strict vs transitional, Word version differences)
- Document known pitfalls and community solutions
- Determine how to handle multiple format versions

**Outcome:** Created `WOTAN/docs/version-compatibility.md` with comprehensive findings on:
- OOXML Strict vs Transitional, namespace extensions (w14-w16sdtdh)
- mc:AlternateContent handling (critical gap in python-docx)
- How docx4j, Open XML SDK, Apache POI handle versions
- Community workarounds and recommendations for WOTAN

---

### B-TEST-01: Baseline Verification and Regression Documentation `[DONE]`

**Intent:** Thoroughly test that python-docx does what it claims to do, document current state, and identify any regressions.

**Next:** T-TEST-01 (completed)

**Details:**
- Run full test suite (pytest + behave) and document results
- Test with example documents in `WOTAN/example-docs/` and `tests/test-files/`
- For each claimed feature, verify it works correctly:
  - Open documents created by Word
  - Read all content types
  - Modify and save
  - Verify round-trip fidelity (open in Word, no corruption)
- Document any failures, regressions, or undocumented limitations
- Create baseline report in `WOTAN/docs/baseline-report.md`
- This establishes ground truth before we add new features

**Acceptance Criteria:**
- All existing tests pass (or failures documented with issues)
- Example documents tested for read/write/round-trip
- Baseline report documents current capabilities and any issues found
- Regression list created if any features don't work as documented

---

## Tier 1: Critical for Full Round-Trip

### B-CORE-01: mc:AlternateContent Handling `[DONE]`

**Intent:** Enable access to content inside `mc:AlternateContent` blocks, which currently is invisible to the API.

**Specification:** `WOTAN/docs/mce-specification.md`

**Research completed:** T-CORE-01 (DONE)

**Details:**
This is the **most critical gap** for the vision of "read any docx, manipulate any part, write it back."

Currently, content inside `mc:AlternateContent` blocks is:
- Invisible to `doc.paragraphs`, `doc.tables`, etc.
- Not accessible through the API
- Preserved in XML on save (so round-trip works at XML level, but not API level)

Affected content includes:
- Text boxes (very common in real documents)
- Modern vector graphics (DrawingML shapes)
- Word 2010+ specific features wrapped for backwards compatibility
- Any content using `wps:wsp`, `wpg:wgrp`, etc.

**Implementation Strategy:** "Choice-First with Preservation"
1. Read from `mc:Choice` (modern, richer content)
2. Write to both `mc:Choice` and `mc:Fallback` (preserve compatibility)
3. Preserve entire structure on round-trip

**Acceptance Criteria:**
- Paragraphs/tables inside mc:AlternateContent are accessible via existing APIs
- Text box content is readable and modifiable
- Round-trip preserves mc:AlternateContent structure
- Document can still be opened in older Word versions (mc:Fallback preserved)

**References:**
- GitHub issue [#1389](https://github.com/python-openxml/python-docx/issues/1389)
- ECMA-376 Part 3 (ISO 29500-3): Markup Compatibility and Extensibility
- `WOTAN/docs/version-compatibility.md` section on MCE

---

### B-DRW-03: Text Box Support `[DONE]`

**Intent:** Provide API access to text boxes, which are a common document element.

**Implementation:** T-CORE-02 (completed as part of B-CORE-01)

**Details:**
Text boxes are now accessible via `doc.text_boxes`:
- `TextBoxes` collection with iteration, indexing, len()
- `TextBox` extends `BlockItemContainer` (paragraphs, tables, add_paragraph, etc.)
- Reads from `mc:Choice` (DrawingML) or `mc:Fallback` (VML)

**Acceptance Criteria:**
- [x] Access text boxes via `doc.text_boxes`
- [x] Read/modify paragraph content inside text boxes
- [ ] Create new text boxes programmatically (deferred to B-DRW-05)
- [ ] Positioning properties (anchor, wrap, size) accessible (deferred)

---

## Tier 2: Write Support for Extensions

### B-SDT-02: Content Control Creation `[READY]`

**Intent:** Enable creating new content controls programmatically.

**Details:**
Current state: Can read content controls and access their content, but cannot create new ones.

Needed:
- `document.add_content_control(type, tag=None, alias=None)`
- Support for text, date, dropdown, comboBox, checkbox types
- Setting dropdown/comboBox list items
- Setting date format and calendar type

**Acceptance Criteria:**
- Create content controls of all supported types
- Set tag, alias, and type-specific properties
- Document opens correctly in Word with functional controls

---

### B-FLD-03: Field Creation `[READY]`

**Intent:** Enable creating new fields programmatically.

**Details:**
Current state: Can read simple and complex fields, but cannot create new ones.

Common use cases:
- Insert PAGE/NUMPAGES fields for page numbering
- Insert DATE fields
- Insert cross-references (REF, PAGEREF)
- Insert HYPERLINK fields

**Acceptance Criteria:**
- Create simple fields with `run.add_field(field_type, switches=None)`
- Create complex fields when needed
- Fields update correctly when document is opened in Word

---

### B-DRW-04: Bookmark Creation `[READY]`

**Intent:** Enable creating new bookmarks programmatically.

**Details:**
Current state: Can read bookmarks, but cannot create new ones.

Needed:
- `document.add_bookmark(name, start_element, end_element=None)` or similar
- Bookmark ID management (must be unique within document)
- Support for point bookmarks (start and end at same location) and range bookmarks

**Acceptance Criteria:**
- Create bookmarks at specific locations
- Unique ID generation/management
- Can be used as targets for cross-references and hyperlinks

---

### B-FN-02: Footnote/Endnote Creation `[READY]`

**Intent:** Enable creating new footnotes and endnotes programmatically.

**Details:**
Current state: Can read footnotes/endnotes and their content, but cannot create new ones.

Needed:
- `run.add_footnote(text=None)` to insert footnote reference and create footnote
- `run.add_endnote(text=None)` for endnotes
- Return footnote/endnote object for adding content

**Acceptance Criteria:**
- Insert footnote/endnote references in text
- Create corresponding footnote/endnote with content
- Proper footnote ID management
- Correct rendering in Word

---

### B-REV-02: Accept/Reject Track Changes `[NEEDS-SPEC]`

**Intent:** Programmatically accept or reject tracked changes.

**Details:**
Current state: Can read track changes (insertions, deletions, author, date), but cannot accept/reject them.

Needed:
- `revision.accept()` - apply the change and remove tracking markup
- `revision.reject()` - revert the change and remove tracking markup
- `document.revisions.accept_all()` / `reject_all()`

**Acceptance Criteria:**
- Accept individual insertions (text becomes normal, w:ins removed)
- Reject individual insertions (text removed)
- Accept individual deletions (text removed, w:del removed)
- Reject individual deletions (text restored)
- Bulk accept/reject all changes

---

## Tier 3: Additional Features

### B-DRW-05: Floating Shape Creation `[READY]`

**Intent:** Enable creating floating (anchored) images and shapes programmatically.

**Details:**
Current state: Can read floating shapes, but cannot create them. Only inline images can be created.

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-DRW-05 section

**Acceptance Criteria:**
- Create anchored images with absolute position
- Set horizontal/vertical alignment
- Set wrap style (square, tight, none, etc.)
- Set behind_text property
- Proper z-order handling

---

### B-STY-02: Theme Modification `[READY]`

**Intent:** Enable modifying theme colors and fonts.

**Details:**
Current state: Can read theme colors and fonts, but cannot modify them.

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-STY-02 section

**Acceptance Criteria:**
- Modify all 12 theme colors
- Modify major and minor fonts
- Changes persist after save
- Document displays correctly in Word

---

### B-CORE-02: Strict OOXML Conformance Detection `[READY]`

**Intent:** Detect and handle OOXML Strict conformance class documents.

**Details:**
Most documents use Transitional conformance, but Strict exists and uses different namespaces.
Currently python-docx assumes Transitional.

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-CORE-02 section

**Acceptance Criteria:**
- Detect Transitional vs Strict conformance
- Detect minimum Word version from namespaces
- Warning when opening Strict documents
- Basic Strict document reading (future: full support)

---

## Completed

### B-TEST-01: Baseline Verification and Regression Documentation `[DONE]`
- All tests pass (1609 pytest, 650 behave scenarios)
- Example documents tested and round-trip verified
- Created `WOTAN/docs/baseline-report.md`
- Identified advanced features in real docs (1299 fields, 348 bookmarks, 63 SDT, etc.)

### B-TEST-02: Research Community Knowledge and Version Compatibility `[DONE]`
- Researched GitHub issues, Stack Overflow, ECMA specs
- Created `WOTAN/docs/version-compatibility.md`
- Key finding: mc:AlternateContent is a critical gap

### B-SDT-01: Content Controls (Structured Document Tags) `[DONE]`
- Created `src/docx/oxml/sdt.py` - OXML classes for SDT elements
- Created `src/docx/sdt.py` - Proxy classes (SdtContentControls, SdtBlockContentControl)
- Added `content_controls` property to Document class
- Supports: text, date, dropdown, comboBox, docPartObj types
- Access by tag/alias, iteration, paragraph/table content

### B-FLD-01: Simple and Complex Fields `[DONE]`
- Created `src/docx/oxml/fields.py` - OXML classes (CT_FldSimple, CT_FldChar, CT_FldInstrText)
- Created `src/docx/fields.py` - Proxy classes (Fields, SimpleField, ComplexFieldProxy)
- Added `fields` property to Document class
- Supports both simple (`w:fldSimple`) and complex (`w:fldChar`) fields
- Filter by type (PAGE, DATE, TOC, REF, CITATION, etc.), access field code and result

### B-NUM-01: Numbering Part Creation `[DONE]`
- Implemented `NumberingPart.new()` in `src/docx/parts/numbering.py`
- Creates minimal numbering.xml with bullet and numbered list definitions
- Enables programmatic list creation for documents without existing numbering part

### B-FN-01: Footnotes and Endnotes `[DONE]`
- Created `src/docx/oxml/footnotes.py` - OXML classes (CT_Footnotes, CT_Footnote, CT_Endnotes, CT_Endnote)
- Created `src/docx/parts/footnotes.py` - FootnotesPart and EndnotesPart
- Created `src/docx/footnotes.py` - Proxy classes (Footnotes, Footnote, Endnotes, Endnote)
- Added `footnotes` and `endnotes` properties to Document class
- Footnotes/endnotes are BlockItemContainer subclasses with full paragraph/table support

### B-DRW-01: Bookmarks `[DONE]`
- Created `src/docx/oxml/bookmarks.py` - OXML classes (CT_Bookmark, CT_MarkupRange)
- Created `src/docx/bookmarks.py` - Proxy classes (Bookmarks, Bookmark)
- Added `bookmarks` property to Document class
- Access by name, iteration, filtering by type (TOC, Ref, Hlk, system vs user)

### B-FLD-02: Complex Fields `[DONE]`
- Already implemented as part of B-FLD-01
- Full support for `w:fldChar` (begin/separate/end) structure
- `w:instrText` field instructions parsing
- ComplexFieldProxy class with field_code, field_type, result properties

### B-REV-01: Track Changes (Read-Only) `[DONE]`
- Created `src/docx/oxml/revisions.py` - OXML classes (CT_RunTrackChange, CT_*PrChange)
- Created `src/docx/revisions.py` - Proxy classes (Revisions, Revision, RevisionType)
- Added `revisions` property to Document class
- Read insertions, deletions with author, date, and text content
- Filter by author, separate access to insertions vs deletions

### B-DRW-02: Floating Images (Anchored Shapes) `[DONE]`
- Extended `CT_Anchor` in `src/docx/oxml/shape.py` with extent, docPr, graphic children
- Created `FloatingShapes` and `FloatingShape` classes in `src/docx/shape.py`
- Added `floating_shapes` property to Document class
- Access width, height, name, description, type, is_behind_text properties

### B-NS-01: Word 2013+ Namespace Support `[DONE]`
- Added namespaces to `src/docx/oxml/ns.py`:
  - `a14` (Office Drawing 2010), `asvg` (SVG 2016)
  - `w15` (Word 2013), `w16` (Word 2016+), `w16se`, `w16cid`, `w16cex`, `w16sdtdh`
  - `wp14` (WordprocessingDrawing 2010)
- Enables xpath queries with modern namespace prefixes
- Preserves modern document features during round-trip

### B-IMG-01: SVG Image Support `[DONE]`
- Added `image/svg+xml` content type to `src/docx/opc/constants.py`
- Created `src/docx/image/svg.py` - SVG image header parser
  - Parses width/height from attributes or viewBox
  - Supports units: px, pt, in, cm, mm, pc, em, ex, %
  - Uses 96 DPI (SVG/CSS standard)
- Added SVG signatures to `src/docx/image/__init__.py`
- Full recognition of SVG images by Image factory

### B-STY-01: Theme Support `[DONE]`
- Created `src/docx/oxml/theme.py` - OXML classes for theme elements
  - CT_OfficeStyleSheet (a:theme), CT_BaseStyles (a:themeElements)
  - CT_ColorScheme (12 theme colors), CT_FontScheme (major/minor fonts)
- Created `src/docx/parts/theme.py` - ThemePart class
- Created `src/docx/theme.py` - Proxy classes
  - Theme (name, colors, fonts), ThemeColors (dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink)
  - ThemeFonts (major_latin, minor_latin, etc.)
- Registered ThemePart in `src/docx/__init__.py`
- Added `theme` property to Document class
- Read theme colors as RGBColor objects, font names as strings
