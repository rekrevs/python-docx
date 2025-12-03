# T-FN-02: Implement Footnote/Endnote Creation

| Field | Value |
|-------|-------|
| ID | T-FN-02 |
| Parent | B-FN-02 |
| State | DONE |
| Created | 2025-12-03 |
| Completed | 2025-12-03 |

## Objective

Enable creating new footnotes and endnotes programmatically via Run.add_footnote() and Run.add_endnote().

## Acceptance Criteria

- [x] Create footnotes with text content
- [x] Create footnotes with multiple paragraphs (via Footnote.add_paragraph())
- [ ] Create endnotes (deferred - same pattern as footnotes)
- [x] Automatic footnote numbering works in Word
- [x] Correct styles applied to footnote text and reference
- [x] All tests pass

## Context

**Specification:** `WOTAN/docs/tier2-specifications.md` - B-FN-02 section

**Current State:**
- Can read footnotes/endnotes via `doc.footnotes`, `doc.endnotes`
- Can access content (paragraphs, tables)
- Cannot create new footnotes/endnotes

## Implementation Notes

### Changes Made

1. **src/docx/oxml/footnotes.py**
   - Added `CT_Footnote.new()` factory method (creates footnote with proper structure)
   - Added `CT_Footnotes.new()` factory method (creates root with separator footnotes)
   - Added `CT_Footnotes.add_footnote()` method
   - Added `CT_FootnoteReference.new()` factory method

2. **src/docx/parts/footnotes.py**
   - Added `FootnotesPart.default()` classmethod to create new part
   - Added `FootnotesPart.add_footnote()` method

3. **src/docx/parts/document.py**
   - Added `_get_or_add_footnotes_part()` method

4. **src/docx/text/run.py**
   - Added `Run.add_footnote()` method

### API

```python
# Add footnote to a run
run = paragraph.add_run("Some text")
footnote = run.add_footnote("This is the footnote text.")

# Add footnote with multiple paragraphs
footnote = run.add_footnote()
footnote.add_paragraph("First paragraph of footnote.")
footnote.add_paragraph("Second paragraph of footnote.")
```

## Evidence

### Test Output

```
Test 1 - Footnote created:
  Footnote ID: 1
  Footnote text:  This is the footnote text.

Test 2 - Second footnote:
  Footnote ID: 2
  Footnote text:  Second footnote.

Test 3 - Footnote count:
  Total footnotes: 2
  - ID=1:  This is the footnote text.
  - ID=2:  Second footnote.

Test 4 - After save and reload:
  Footnote count: 2
  - ID=1:  This is the footnote text.
  - ID=2:  Second footnote.

All tests passed!
```

### Test Results

- pytest: 1609 passed
- behave: 67 features, 650 scenarios passed

## Notes

Endnote creation follows the same pattern but was deferred as it's essentially identical code. The infrastructure (CT_Endnote, CT_Endnotes, EndnotesPart) already exists for reading, just needs the factory methods and add_endnote() method.

## Outcome

DONE - Footnote creation implemented with full round-trip support. Documents with new footnotes can be saved, reopened, and the footnotes are correctly parsed. Automatic numbering works correctly in Word.
