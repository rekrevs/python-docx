# T-TEST-01: Baseline Verification and Regression Documentation

| Field | Value |
|-------|-------|
| ID | T-TEST-01 |
| Parent | B-TEST-01 |
| State | DONE |
| Created | 2024-12-03 |
| Completed | 2024-12-03 |

## Objective

Thoroughly verify that python-docx works as documented, run all tests, test with real documents, and create a baseline report of the current state including any regressions or issues found.

## Acceptance Criteria

- [x] All pytest unit tests pass (or failures documented)
- [x] All behave acceptance tests pass (or failures documented)
- [x] Type checking (pyright) passes — Has errors but pre-existing, documented
- [x] Linting (ruff) passes
- [x] Test documents from `tests/test-files/` load correctly
- [x] Example documents from `WOTAN/example-docs/` tested
- [x] Round-trip test: open, modify, save, verify no corruption
- [x] Baseline report created at `WOTAN/docs/baseline-report.md`

## Context

Before adding new features, we need to establish that the existing codebase works correctly. This task:
1. Runs the full test suite
2. Tests with real Word documents
3. Documents any issues found
4. Creates a baseline for future regression testing

## Test Files

### Existing test files
- `tests/test_files/` - 4 docx files, all pass

### Example documents
- `WOTAN/example-docs/` - 4 real-world Word documents tested

## Subtasks

| ID | Description | State |
|----|-------------|-------|
| T-TEST-01-1 | Run pytest and document results | DONE |
| T-TEST-01-2 | Run behave and document results | DONE |
| T-TEST-01-3 | Run pyright and ruff | DONE |
| T-TEST-01-4 | Test existing test-files | DONE |
| T-TEST-01-5 | Test WOTAN/example-docs | DONE |
| T-TEST-01-6 | Create baseline report | DONE |

## Implementation Notes

### Test Suite Results
- **pytest:** 1609 tests passed in 2.20s
- **behave:** 67 features, 650 scenarios, 1856 steps — all passed
- **ruff:** All checks passed
- **pyright:** 2438 errors in src/ (pre-existing, due to xmlchemy dynamic types)

### Document Testing
All 8 documents (4 test + 4 example) load and round-trip successfully.

### Advanced Feature Discovery
Real-world documents contain features python-docx cannot access:
- 63 Content Controls (SDT)
- 1299 Complex Fields
- 348 Bookmarks
- 14 Footnotes
- 43 Track Changes

These features are **preserved** during round-trip but **not accessible** via API.

## Obstacles

None — testing proceeded smoothly.

## Evidence

### pytest output
```
============================= 1609 passed in 2.20s =============================
```

### behave output
```
67 features passed, 0 failed, 0 skipped
650 scenarios passed, 0 failed, 0 skipped
1856 steps passed, 0 failed, 0 skipped
```

### Document load test
```
OK blk-inner-content.docx - Paragraphs: 2, Tables: 1, Sections: 1
OK having-images.docx - Paragraphs: 5, Tables: 0, Sections: 1
OK sct-inner-content.docx - Paragraphs: 7, Tables: 2, Sections: 3
OK test.docx - Paragraphs: 2, Tables: 0, Sections: 1
OK Ang. examensarbete.docx - Paragraphs: 130
OK Q-NEXUS_Application Form.docx - Paragraphs: 343, Tables: 23, Comments: 66
OK RE20221684 Nautisk Riskanalys.docx - Paragraphs: 829, Tables: 27, Images: 86
OK SK25 Datorsystem project plan.docx - Paragraphs: 75
```

### Round-trip test
All 8 documents: open → add paragraph → save → reopen = success

## Outcome

**DONE**

python-docx 1.2.0 is stable and passes all tests. No regressions identified.

Key finding: Real-world documents contain many advanced features (SDT, fields, bookmarks, footnotes, track changes) that python-docx preserves but cannot access. This validates the WOTAN backlog priorities.

Full details in `WOTAN/docs/baseline-report.md`.
