# T-CORE-03: Implement Strict OOXML Conformance Detection

| Field | Value |
|-------|-------|
| ID | T-CORE-03 |
| Parent | B-CORE-02 |
| State | DONE |
| Created | 2025-12-03 |
| Completed | 2025-12-03 |

## Objective

Detect and report OOXML conformance class (Transitional vs Strict) and minimum Word version required for a document.

## Acceptance Criteria

- [x] Detect Transitional vs Strict conformance
- [x] Detect minimum Word version from namespaces
- [x] List supported namespaces
- [ ] Warning when opening Strict documents (deferred)
- [ ] Basic Strict document reading (future: full support - deferred)

## Context

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-CORE-02 section

**Current State:**
- python-docx assumes Transitional conformance
- No detection of Strict documents
- No version information available

**Namespace Differences:**
- Transitional: `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- Strict: `http://purl.oclc.org/ooxml/wordprocessingml/main`

**Word Version Detection via mc:Ignorable:**
- No w14: Word 2007
- w14: Word 2010+
- w15: Word 2013+
- w16: Word 2016+
- w16cid, w16cex: Word 365/2019+
- w16sdtdh: Word 2021+

## Implementation Notes

### Changes Made

Added three new properties to `Document` class in `src/docx/document.py`:

1. **`document.conformance`** - Returns 'transitional' or 'strict'
   - Checks the 'w' namespace in the document element's nsmap
   - Strict uses `purl.oclc.org` in the namespace URI

2. **`document.minimum_word_version`** - Returns version string
   - Parses `mc:Ignorable` attribute
   - Returns one of: 'Word 2007', 'Word 2010', 'Word 2013', 'Word 2016', 'Word 365/2019', 'Word 2021+'

3. **`document.supported_namespaces`** - Returns list of namespace prefixes
   - Extracts from `mc:Ignorable` attribute
   - Common prefixes: w14, w15, w16, w16se, w16cid, w16cex, w16sdtdh, wp14

### API

```python
from docx import Document

doc = Document('document.docx')

# Check conformance class
print(f'Conformance: {doc.conformance}')  # 'transitional' or 'strict'

# Check minimum Word version
print(f'Requires: {doc.minimum_word_version}')  # e.g., 'Word 2021+'

# List supported namespaces
print(f'Namespaces: {doc.supported_namespaces}')
# e.g., ['w14', 'w15', 'w16se', 'w16cid', 'w16', 'w16cex', 'w16sdtdh', 'wp14']
```

## Obstacles

None encountered.

## Evidence

### Test Output

```
Default document:
  Conformance: transitional
  Min Word version: Word 2010
  Supported namespaces: ['w14', 'wp14']

Real document:
  Conformance: transitional
  Min Word version: Word 2021+
  Supported namespaces: ['w14', 'w15', 'w16se', 'w16cid', 'w16', 'w16cex', 'w16sdtdh', 'w16sdtfl', 'w16du', 'wp14']

All tests passed!
```

### Test Results

- pytest: 1609 passed
- behave: 67 features, 650 scenarios passed

## Outcome

DONE - Conformance detection implemented. Three new Document properties allow inspection of:
- OOXML conformance class (transitional/strict)
- Minimum Word version required
- List of extension namespaces used

Note: Warning on opening Strict documents and actual Strict document reading support are deferred for future work as Strict documents are rare.
