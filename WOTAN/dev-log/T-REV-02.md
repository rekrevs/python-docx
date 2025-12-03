# T-REV-02: Implement Accept/Reject Track Changes

| Field | Value |
|-------|-------|
| ID | T-REV-02 |
| Parent | B-REV-02 |
| State | DONE |
| Created | 2025-12-03 |
| Completed | 2025-12-03 |

## Objective

Enable programmatically accepting or rejecting tracked changes (revisions).

## Acceptance Criteria

- [x] Accept individual insertions (text becomes normal, w:ins removed)
- [x] Reject individual insertions (text removed)
- [x] Accept individual deletions (text removed, w:del removed)
- [x] Reject individual deletions (text restored with w:delText converted to w:t)
- [x] Bulk accept_all() / reject_all()
- [x] All tests pass

## Context

**Specification:** `WOTAN/backlog.md` - B-REV-02 section

**Current State:**
- Can read track changes via `doc.revisions`
- Can access text, author, date, revision_id
- Cannot accept or reject changes

**XML Structure:**

Insertion:
```xml
<w:ins w:id="1" w:author="John" w:date="2025-01-15T10:30:00Z">
    <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

Deletion:
```xml
<w:del w:id="2" w:author="Jane" w:date="2025-01-15T11:00:00Z">
    <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Accept/Reject Logic:**

1. **Accept insertion**: Move runs outside `w:ins`, remove `w:ins`
2. **Reject insertion**: Remove entire `w:ins` element
3. **Accept deletion**: Remove entire `w:del` element (text stays deleted)
4. **Reject deletion**: Convert `w:delText` to `w:t`, move runs outside `w:del`, remove `w:del`

## Implementation Notes

### Changes Made

1. **src/docx/oxml/revisions.py**
   - Added `CT_RunTrackChange.is_insertion` and `is_deletion` properties
   - Added `CT_RunTrackChange.accept()` method
   - Added `CT_RunTrackChange.reject()` method
   - Renamed `text` property to `revision_text` to avoid lxml conflict

2. **src/docx/revisions.py**
   - Added `Revision.accept()` method
   - Added `Revision.reject()` method
   - Added `Revisions.accept_all()` method
   - Added `Revisions.reject_all()` method

### API

```python
# Accept individual revision
for revision in document.revisions:
    if revision.author == "John":
        revision.accept()

# Reject all deletions
for revision in document.revisions.deletions:
    revision.reject()

# Bulk operations
count = document.revisions.accept_all()  # Returns count accepted
count = document.revisions.reject_all()  # Returns count rejected
```

## Evidence

### Test Output

```
Test 1 - Reject insertion:
  Before: "Start.  Middle.  End."
  After:  "Start.  Middle.  End."
  PASSED

Test 2 - Accept deletion:
  Before: "Start.  Middle.  End." (DELETED is tracked)
  After:  "Start.  Middle.  End." (DELETED stays removed)
  PASSED

Test 3 - accept_all():
  Revisions before: 2
  Accepted: 2
  Revisions after: 0
  Text: "Start. INSERTED Middle.  End."
  PASSED

Test 4 - reject_all():
  Revisions before: 2
  Rejected: 2
  Revisions after: 0
  Text: "Start.  Middle. DELETED End."
  PASSED

Test 5 - Round-trip:
  Revisions after reload: 0
  Text: "Start. INSERTED Middle.  End."
  PASSED

=== All tests passed! ===
```

### Test Results

- pytest: 1609 passed
- behave: 67 features, 650 scenarios passed

## Outcome

DONE - Accept/reject track changes implemented with full functionality. Individual revisions can be accepted or rejected, and bulk operations are available via accept_all() and reject_all().
