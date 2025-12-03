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

### B-SDT-02: Content Control Creation `[DONE]`

**Intent:** Enable creating new content controls programmatically.

**Next:** T-SDT-02 (completed)

**Details:**
Implemented via `document.add_content_control(type, tag=None, alias=None)`:
- Support for text, date, dropdown, comboBox types
- Setting tag and alias properties
- Block-level content controls only (for now)

**Acceptance Criteria:**
- [x] Create content controls of supported types
- [x] Set tag, alias, and type-specific properties
- [x] Document opens correctly in Word with functional controls

---

### B-FLD-03: Field Creation `[DONE]`

**Intent:** Enable creating new fields programmatically.

**Next:** T-FLD-03 (completed)

**Details:**
Implemented both simple and complex field creation:
- `run.add_simple_field(field_code)` for simple fields
- `run.add_complex_field(field_code)` for complex fields
- Support for PAGE, NUMPAGES, DATE, REF, HYPERLINK, etc.

**Acceptance Criteria:**
- [x] Create simple fields
- [x] Create complex fields when needed
- [x] Fields update correctly when document is opened in Word

---

### B-DRW-04: Bookmark Creation `[DONE]`

**Intent:** Enable creating new bookmarks programmatically.

**Next:** T-DRW-04 (completed)

**Details:**
Implemented `paragraph.add_bookmark(name)` and `run.add_bookmark(name)`:
- Automatic unique ID generation
- Point bookmarks (start and end at same location)

**Acceptance Criteria:**
- [x] Create bookmarks at specific locations
- [x] Unique ID generation/management
- [x] Can be used as targets for cross-references and hyperlinks

---

### B-FN-02: Footnote/Endnote Creation `[DONE]`

**Intent:** Enable creating new footnotes and endnotes programmatically.

**Next:** T-FN-02 (completed)

**Details:**
Implemented `run.add_footnote_reference()` and `run.add_endnote_reference()`:
- Inserts reference mark in text
- Creates footnote/endnote in respective part
- Returns footnote/endnote object for adding content
- Automatic ID management

**Acceptance Criteria:**
- [x] Insert footnote/endnote references in text
- [x] Create corresponding footnote/endnote with content
- [x] Proper footnote ID management
- [x] Correct rendering in Word

---

### B-REV-02: Accept/Reject Track Changes `[DONE]`

**Intent:** Programmatically accept or reject tracked changes.

**Next:** T-REV-02 (completed)

**Details:**
Implemented `revision.accept()`, `revision.reject()`, and bulk methods:
- Accept insertions: moves runs out of w:ins, removes wrapper
- Reject insertions: removes entire w:ins with contents
- Accept deletions: removes w:del element entirely
- Reject deletions: converts w:delText back to w:t

**Acceptance Criteria:**
- [x] Accept individual insertions (text becomes normal, w:ins removed)
- [x] Reject individual insertions (text removed)
- [x] Accept individual deletions (text removed, w:del removed)
- [x] Reject individual deletions (text restored)
- [x] Bulk accept/reject all changes

---

## Tier 3: Additional Features

### B-DRW-05: Floating Shape Creation `[DONE]`

**Intent:** Enable creating floating (anchored) images and shapes programmatically.

**Next:** T-DRW-05 (completed)

**Details:**
Implemented `document.add_floating_picture()`:
- Create anchored images with absolute position
- Support wrap types: none, square, tight, through, topAndBottom
- Support behind_doc, h_relative_from, v_relative_from

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-DRW-05 section

**Acceptance Criteria:**
- [x] Create anchored images with absolute position
- [x] Set horizontal/vertical alignment (via relative_from)
- [x] Set wrap style (square, tight, none, etc.)
- [x] Set behind_text property
- [x] Proper z-order handling

---

### B-STY-02: Theme Modification `[DONE]`

**Intent:** Enable modifying theme colors and fonts.

**Next:** T-STY-02 (completed)

**Details:**
Implemented setters for theme colors and fonts:
- `doc.theme.colors.accent1 = RGBColor(...)` or hex string
- `doc.theme.fonts.major_latin = "Arial"`
- All 12 colors and 6 font slots modifiable

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-STY-02 section

**Acceptance Criteria:**
- [x] Modify all 12 theme colors
- [x] Modify major and minor fonts
- [x] Changes persist after save
- [x] Document displays correctly in Word

---

### B-CORE-02: Strict OOXML Conformance Detection `[DONE]`

**Intent:** Detect and handle OOXML Strict conformance class documents.

**Next:** T-CORE-03 (completed)

**Details:**
Implemented three new properties on Document class:
- `doc.conformance` - Returns 'transitional' or 'strict'
- `doc.minimum_word_version` - Returns version string (Word 2007 through Word 2021+)
- `doc.supported_namespaces` - Returns list of namespace prefixes

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-CORE-02 section

**Acceptance Criteria:**
- [x] Detect Transitional vs Strict conformance
- [x] Detect minimum Word version from namespaces
- [x] List supported extension namespaces
- [ ] Warning when opening Strict documents (deferred)
- [ ] Basic Strict document reading (deferred - Strict documents are rare)

---

## Tier 4: Remaining Gaps

### B-FLD-04: Field Modification `[READY]`

**Intent:** Enable modifying existing fields.

**Details:**
Current state: Can read fields and create new ones, but cannot modify existing fields.

Needed:
- `field.field_code = "NEW CODE"` - Change field instruction
- `field.update()` - Trigger field update (may require Word)
- `field.delete()` - Remove field from document
- `field.convert_to_text()` - Replace field with its current result

**Acceptance Criteria:**
- Modify field code of simple fields
- Delete fields (both simple and complex)
- Convert field to static text
- Handle complex field modification (begin/separate/end structure)

---

### B-DRW-08: Floating Shape Modification `[READY]`

**Intent:** Enable modifying properties of existing floating shapes.

**Details:**
Current state: Can read floating shape properties and create new ones, but all properties are read-only.

Needed:
- `shape.width = Inches(2)` - Resize width
- `shape.height = Inches(3)` - Resize height
- `shape.pos_x = Inches(1)` - Move horizontally
- `shape.pos_y = Inches(2)` - Move vertically
- `shape.name = "New Name"` - Rename shape
- `shape.wrap_type = "tight"` - Change text wrapping
- `shape.is_behind_text = True` - Move behind/in front of text
- `shape.delete()` - Remove shape from document

**Acceptance Criteria:**
- Resize floating shapes
- Reposition floating shapes
- Change wrap style
- Change z-order (behind/in front)
- Delete floating shapes

---

### B-DRW-06: Text Box Creation `[READY]`

**Intent:** Enable creating new text boxes programmatically.

**Details:**
Current state: Can read text boxes via `doc.text_boxes`, but cannot create new ones.

Text boxes require:
- `mc:AlternateContent` wrapper with Choice (DrawingML) and Fallback (VML)
- `wp:anchor` for positioning
- `wps:wsp` (WordprocessingShape) containing `wps:txbx` with `w:txbxContent`

**Acceptance Criteria:**
- Create text boxes with specified position and size
- Add paragraphs and tables to text box content
- Set border, fill, and text wrapping properties
- Round-trip preserves structure

---

### B-DRW-07: Bookmark Modification `[READY]`

**Intent:** Enable moving and deleting bookmarks.

**Details:**
Current state: Can read bookmarks and create new ones, but cannot modify existing bookmarks.

Needed:
- `bookmark.delete()` - Remove bookmark from document
- `bookmark.move_to(element)` - Move bookmark to new location
- `bookmark.set_range(start, end)` - Change bookmark range

**Acceptance Criteria:**
- Delete bookmarks by name or object
- Move bookmark start/end to new locations
- Change bookmark from point to range or vice versa

---

### B-MATH-01: Math Equations `[NEEDS-SPEC]`

**Intent:** Read and create mathematical equations (OMML).

**Details:**
Office Math Markup Language (OMML) is used for equations in Word documents.
Located in `m:oMath` and `m:oMathPara` elements.

Common equation elements:
- `m:r` - Math run (text)
- `m:f` - Fraction
- `m:rad` - Radical (square root)
- `m:sSup`, `m:sSub` - Superscript/subscript
- `m:nary` - N-ary operator (sum, integral)
- `m:m` - Matrix
- `m:d` - Delimiter (parentheses, brackets)

**Acceptance Criteria:**
- Read equation content as structured objects
- Access equation components (fractions, radicals, etc.)
- Create simple equations programmatically
- Round-trip preserves equation structure

**References:**
- ECMA-376 Part 1, Section 22 (Office Math)
- [OMML documentation](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/)

---

### B-CHART-01: Chart Support `[NEEDS-SPEC]`

**Intent:** Read and manipulate charts embedded in documents.

**Details:**
Charts in Word are DrawingML charts stored in separate parts (`/word/charts/chartN.xml`).
Referenced via `c:chart` in drawing elements.

Chart structure:
- `c:chartSpace` - Root element
- `c:chart` - Chart definition
- `c:plotArea` - Plot area with axes and series
- `c:ser` - Data series

**Acceptance Criteria:**
- Detect charts in document
- Read chart type (bar, line, pie, etc.)
- Access chart data series and values
- Modify chart data
- (Future) Create simple charts

**References:**
- ECMA-376 Part 1, Section 21.2 (DrawingML Charts)

---

### B-SMART-01: SmartArt Support `[NEEDS-SPEC]`

**Intent:** Read SmartArt diagrams.

**Details:**
SmartArt is stored as DrawingML diagrams in separate parts.
Complex structure with layout, data, and drawing components.

Parts involved:
- `/word/diagrams/data.xml` - Diagram data
- `/word/diagrams/layout.xml` - Layout definition
- `/word/diagrams/quickStyle.xml` - Quick style
- `/word/diagrams/colors.xml` - Color scheme
- `/word/diagrams/drawing.xml` - Visual representation

**Acceptance Criteria:**
- Detect SmartArt in document
- Read text content from SmartArt nodes
- Identify SmartArt type/layout
- (Future) Modify SmartArt text content

**References:**
- ECMA-376 Part 1, Section 21.4 (DrawingML Diagrams)

---

### B-XML-01: Custom XML Support `[NEEDS-SPEC]`

**Intent:** Read and manipulate Custom XML parts and data bindings.

**Details:**
Documents can contain custom XML data in:
- Custom XML parts (`/customXml/itemN.xml`)
- Custom XML data stores
- Content controls bound to custom XML via `w:dataBinding`

Used for:
- Data-driven document generation
- Document metadata
- Integration with external systems

**Acceptance Criteria:**
- Read custom XML parts
- Access custom XML data by XPath
- Modify custom XML values
- Create new custom XML parts
- Understand SDT data bindings to custom XML

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
