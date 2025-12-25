# Implementation Status

This document tracks the current implementation status of python-docx features and WOTAN extensions.

**Last updated:** 2025-12-03

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
| Table borders | Basic | Limited style options |

---

## WOTAN Extensions

### Fully Implemented

| Feature | Read | Modify | Create | Notes |
|---------|:----:|:------:|:------:|-------|
| **Content Controls (SDT)** | ✅ | ✅ | ✅ | text, richText, date, dropDownList, comboBox |
| **Simple Fields** | ✅ | ✅ | ✅ | PAGE, DATE, etc. with delete/convert_to_text |
| **Complex Fields** | ✅ | ✅ | ✅ | TOC, REF, HYPERLINK with delete/convert_to_text |
| **Footnotes** | ✅ | ✅ | ✅ | Full paragraph/table support |
| **Endnotes** | ✅ | ✅ | ✅ | Full paragraph/table support |
| **Bookmarks** | ✅ | ✅ | ✅ | Rename, delete, create |
| **Track Changes** | ✅ | ✅ | — | Accept/reject individual or all |
| **Floating Shapes** | ✅ | ✅ | ✅ | Resize, reposition, rename, delete |
| **Text Boxes** | ✅ | ✅ | ✅ | mc:AlternateContent with Choice/Fallback |
| **Theme Colors** | ✅ | ✅ | — | All 12 theme colors |
| **Theme Fonts** | ✅ | ✅ | — | Major/minor latin, east asian, complex |
| **Math Equations** | ✅ | ✅ | ✅ | OMML iteration and creation |
| **Charts** | ✅ | — | — | Detection and name access |
| **SmartArt** | ✅ | — | — | Detection and name access |
| **Custom XML** | ✅ | ✅ | ✅ | Full part CRUD |
| **SVG Images** | ✅ | — | — | Detection and parsing |
| **Conformance Detection** | ✅ | n/a | n/a | Strict vs Transitional |
| **Word Version Detection** | ✅ | n/a | n/a | Word 2007 through 2021+ |
| **Modern Namespaces** | ✅ | n/a | n/a | w14, w15, w16, w16cex, w16cid, etc. |

### Legend

- ✅ = Full support
- — = Not applicable or not yet implemented

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
| Word 2010 | `w14` | Full |
| Word 2013 | `w15` | Full |
| Word 2018 | `w16` | Full |
| Word 2020 extensions | `w16cex`, `w16cid`, `w16sdtdh`, `w16se` | Full |
| DrawingML | `a`, `a14` | Full |
| Pictures | `pic` | Full |
| WordprocessingDrawing | `wp`, `wp14` | Full |
| WordprocessingShape | `wps` | Full |
| Charts | `c` | Detection |
| Diagrams | `dgm` | Detection |
| Math | `m` | Full |
| Relationships | `r` | Full |
| Custom XML | `ds` | Full |
| VML | `v`, `o` | Full |
| SVG | `asvg` | Full |
| Markup Compatibility | `mc` | Full |
