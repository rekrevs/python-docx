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

## Tier 1: High Priority

### B-SDT-01: Content Controls (Structured Document Tags) `[READY]`

**Intent:** Implement support for `<w:sdt>` elements to enable form creation and template automation.

**Next:** None yet

**Details:**
- Content controls are critical for fillable forms and data binding
- Need to support: plain text, rich text, dropdown, date picker, checkbox types
- Requires new OXML classes: CT_SdtBlock, CT_SdtRun, CT_SdtPr, CT_SdtContent
- Should integrate with existing paragraph/run iteration
- Reference: ECMA-376 Part 1, Section 17.5.2

---

### B-FLD-01: Simple Fields `[READY]`

**Intent:** Implement support for simple field codes (PAGE, DATE, NUMPAGES, etc.)

**Next:** None yet

**Details:**
- Start with `<w:fldSimple>` element for basic fields
- Later extend to complex fields with `<w:fldChar>` and `<w:instrText>`
- Common use cases: page numbers, dates, document properties
- Reference: ECMA-376 Part 1, Section 17.16

---

### B-NUM-01: Numbering Part Creation `[READY]`

**Intent:** Fix `NumberingPart.new()` to allow programmatic list creation.

**Next:** None yet

**Details:**
- Currently raises `NotImplementedError`
- Need to create abstract numbering definitions
- Allow defining bullet and numbered list styles from scratch
- Foundation exists in `oxml/numbering.py` but incomplete
- Reference: ECMA-376 Part 1, Section 17.9

---

### B-FN-01: Footnotes and Endnotes `[READY]`

**Intent:** Implement support for footnotes and endnotes.

**Next:** None yet

**Details:**
- Content types already defined in OPC constants
- Need: FootnotesPart, EndnotesPart
- Need OXML: CT_Footnotes, CT_Footnote, CT_FtnEdnRef
- API: paragraph.add_footnote(), document.footnotes collection
- Reference: ECMA-376 Part 1, Section 17.11

---

## Tier 2: Medium Priority

### B-DRW-01: Bookmarks `[READY]`

**Intent:** Implement bookmark support for named locations and cross-references.

**Next:** None yet

**Details:**
- Need: `<w:bookmarkStart>`, `<w:bookmarkEnd>` elements
- API: document.bookmarks collection, paragraph.add_bookmark()
- Enable cross-references to bookmarks
- Relatively straightforward XML structure
- Reference: ECMA-376 Part 1, Section 17.13.6

---

### B-REV-01: Track Changes (Read-Only) `[NEEDS-SPEC]`

**Intent:** Add ability to detect and read revision marks.

**Next:** None yet

**Details:**
- Start with read-only support for `<w:ins>`, `<w:del>`, `<w:rPrChange>`
- Expose revision metadata (author, date)
- Later consider accept/reject functionality
- Complex feature with many element types
- Reference: ECMA-376 Part 1, Section 17.13.5

---

### B-DRW-02: Floating Images (Anchored Shapes) `[NEEDS-SPEC]`

**Intent:** Extend drawing support to handle `wp:anchor` positioned graphics.

**Next:** None yet

**Details:**
- Currently only `wp:inline` pictures supported
- Need anchor positioning, text wrapping options
- Extend InlineShape or create new FloatingShape class
- Reference: ECMA-376 Part 1, DrawingML sections

---

### B-FLD-02: Complex Fields `[BLOCKED]`

**Intent:** Support complex field codes (TOC, cross-refs, MERGEFIELD).

**Next:** Depends on B-FLD-01

**Details:**
- Requires `<w:fldChar>` (begin/separate/end) structure
- `<w:instrText>` for field instructions
- More complex than simple fields
- Reference: ECMA-376 Part 1, Section 17.16

---

## Tier 3: Lower Priority

### B-NS-01: Word 2013+ Namespace Support `[NEEDS-SPEC]`

**Intent:** Add missing namespace declarations for modern Word features.

**Next:** None yet

**Details:**
- Add to ns.py: w15, w16, w16se, w16cid, w16cex, w16sdtdh, wp14
- Enables reading/preserving modern document features
- Required for proper round-tripping of Office 365 documents

---

### B-IMG-01: SVG Image Support `[NEEDS-SPEC]`

**Intent:** Add support for SVG images in documents.

**Next:** None yet

**Details:**
- Common in modern documents
- Need SVG image header parser
- May require special handling for fallback PNG

---

### B-STY-01: Theme Support `[NEEDS-SPEC]`

**Intent:** Implement document theme reading and modification.

**Next:** None yet

**Details:**
- Theme affects colors, fonts, effects cascade
- ThemePart exists in OPC constants but no API
- Complex color scheme resolution

---

## Completed

(None yet)
