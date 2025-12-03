# Implementation Status

This document tracks the current implementation status of python-docx features and WOTAN extensions.

## Core python-docx (Upstream)

### Fully Implemented

| Feature | Version | Notes |
|---------|---------|-------|
| Document creation/opening | 0.8+ | |
| Document saving | 0.8+ | |
| Paragraphs | 0.8+ | Full CRUD |
| Runs | 0.8+ | Full CRUD |
| Tables | 0.8+ | Create, modify, merge cells |
| Styles | 0.8+ | Paragraph, character, table |
| Headers/Footers | 0.8+ | Even/odd, first page, sections |
| Sections | 0.8+ | Page setup, orientation, margins |
| Inline images | 0.8+ | PNG, JPEG, TIFF, BMP, GIF |
| Hyperlinks | 1.0+ | External links |
| Core properties | 0.8+ | Title, author, etc. |
| Comments | 1.2.0 | Full support with ranges |
| Character formatting | 0.8+ | Bold, italic, underline, color, etc. |
| Paragraph formatting | 0.8+ | Alignment, spacing, indentation |
| Tab stops | 0.8+ | With leaders |
| Breaks | 0.8+ | Page, line, column |

### Partially Implemented

| Feature | Status | Limitation |
|---------|--------|------------|
| Numbering | Read-only | Cannot create new definitions |
| Settings | Minimal | Only `evenAndOddHeaders` exposed |
| Floating images | Registered | `wp:anchor` not functional |
| Table borders | Basic | Limited style options |

### Not Implemented

| Feature | Priority | Backlog Item |
|---------|----------|--------------|
| Content controls (SDT) | High | B-SDT-01 |
| Fields | High | B-FLD-01 |
| Footnotes/Endnotes | High | B-FN-01 |
| Bookmarks | Medium | B-DRW-01 |
| Track changes | Medium | B-REV-01 |
| Math equations | Low | - |
| Charts | Low | - |
| SmartArt | Low | - |
| Themes | Low | - |

---

## WOTAN Extensions

### In Progress

(None yet)

### Completed

(None yet)

---

## Test Coverage

### Unit Tests (`tests/`)

| Module | Coverage | Notes |
|--------|----------|-------|
| document | Good | |
| paragraph | Good | |
| run | Good | |
| table | Good | |
| section | Good | |
| styles | Good | |
| comments | Good | New in 1.2.0 |
| oxml/* | Good | |
| opc/* | Good | |

### Acceptance Tests (`features/`)

71 feature files covering:
- API and document operations
- Block-level items
- Comments
- Headers/footers
- Hyperlinks
- Images
- Numbering (read)
- Paragraphs
- Runs
- Sections
- Styles
- Tables
- Text formatting

---

## Version History

| Version | Date | Notable Features |
|---------|------|------------------|
| 1.2.0 | 2024 | Comments support |
| 1.1.0 | 2023 | Type hints, Python 3.9+ |
| 1.0.0 | 2022 | Stable API |
| 0.8.x | 2018-2021 | Core features |

---

## Namespace Support

| Namespace | Prefix | Status |
|-----------|--------|--------|
| WordprocessingML 2006 | `w` | Full |
| Word 2010 | `w14` | Defined |
| Word 2013 | `w15` | Not defined |
| Word 2018 | `w16` | Not defined |
| DrawingML | `a` | Partial |
| Pictures | `pic` | Full |
| WordprocessingDrawing | `wp` | Partial |
| Charts | `c` | Detection only |
| Diagrams | `dgm` | Detection only |
| Math | `m` | Preserved |
| Relationships | `r` | Full |
