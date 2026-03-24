# Gap Analysis: python-docx vs OOXML Specification

This document captures the comprehensive analysis of what python-docx currently supports versus what the full OOXML (ECMA-376 / ISO/IEC 29500) specification provides.

## 1. Specification Overview

### Standards Landscape

| Standard | Version | Status |
|----------|---------|--------|
| ECMA-376 | 5th Edition (Dec 2016) | Current |
| ISO/IEC 29500-1 | 2016 | Fundamentals & Markup |
| ISO/IEC 29500-2 | 2021 | Open Packaging Conventions |
| ISO/IEC 29500-3 | 2015 | Markup Compatibility |
| ISO/IEC 29500-4 | 2016 | Transitional Features |
| MS-DOCX | v22.1 (Nov 2025) | Word Extensions |

### Microsoft Extension Namespaces

| Prefix | URI | Word Version | python-docx |
|--------|-----|--------------|-------------|
| `w` | `.../wordprocessingml/2006/main` | All | Supported |
| `w14` | `.../word/2010/wordml` | 2010 | Defined |
| `w15` | `.../word/2012/wordml` | 2013 | Defined |
| `w16` | `.../word/2018/wordml` | 2018 | Defined |
| `w16se` | `.../word/2015/wordml/symex` | 2015 | Defined |
| `w16cid` | `.../word/2016/wordml/cid` | 2016 | Defined |
| `w16cex` | `.../word/2018/wordml/cex` | 2018 | Defined |
| `w16sdtdh` | `.../word/2020/wordml/sdtdatahash` | 2020 | Defined |
| `w16du` | `.../word/2023/wordml/word16du` | 2023 | Defined (T-0015) |
| `w16sfl` | `.../word/2024/wordml/sdtformatlock` | 2024 | Defined (T-0015) |
| `wp14` | `.../word/2010/wordprocessingDrawing` | 2010 | Defined |

---

## 2. Feature Support Matrix

### Legend
- ✅ Full support
- ⚠️ Partial support
- ❌ Not implemented
- 🔍 Read-only / Detection only

### 2.1 Document Structure

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Document body | `w:document`, `w:body` | ✅ | |
| Paragraphs | `w:p` | ✅ | |
| Runs | `w:r` | ✅ | |
| Text | `w:t` | ✅ | |
| Sections | `w:sectPr` | ✅ | |
| Headers | `w:hdr` | ✅ | Even/odd, first page |
| Footers | `w:ftr` | ✅ | Even/odd, first page |
| Body sectPr | `w:sectPr` in body | ✅ | |

### 2.2 Block-Level Content

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Paragraphs | `w:p` | ✅ | |
| Tables | `w:tbl` | ✅ | |
| Content controls (block) | `w:sdt` | ❌ | High priority |
| Custom XML (block) | `w:customXml` | ❌ | |
| Alternate content | `mc:AlternateContent` | ❌ | |

### 2.3 Inline Content

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Runs | `w:r` | ✅ | |
| Hyperlinks | `w:hyperlink` | ✅ | |
| Inline images | `w:drawing/wp:inline` | ✅ | |
| Floating images | `w:drawing/wp:anchor` | ⚠️ | Minimal |
| Simple fields | `w:fldSimple` | ❌ | High priority |
| Complex fields | `w:fldChar`, `w:instrText` | ❌ | |
| Bookmarks | `w:bookmarkStart/End` | ❌ | |
| Comments | `w:commentRangeStart/End` | ✅ | v1.2.0 |
| Footnote refs | `w:footnoteReference` | ❌ | |
| Endnote refs | `w:endnoteReference` | ❌ | |
| Content controls (inline) | `w:sdt` | ❌ | |
| Ruby (East Asian) | `w:ruby` | ❌ | |
| Symbols | `w:sym` | ❌ | |

### 2.4 Tables

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Table structure | `w:tbl`, `w:tr`, `w:tc` | ✅ | |
| Table properties | `w:tblPr` | ✅ | |
| Row properties | `w:trPr` | ✅ | |
| Cell properties | `w:tcPr` | ✅ | |
| Cell merging (horizontal) | `w:gridSpan` | ✅ | |
| Cell merging (vertical) | `w:vMerge` | ✅ | |
| Table styles | `w:tblStyle` | ✅ | |
| Nested tables | | ✅ | |
| Table grid | `w:tblGrid` | ✅ | |

### 2.5 Character Formatting

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Bold | `w:b` | ✅ | |
| Italic | `w:i` | ✅ | |
| Underline | `w:u` | ✅ | |
| Strikethrough | `w:strike` | ✅ | |
| Double strikethrough | `w:dstrike` | ✅ | |
| Font name | `w:rFonts` | ✅ | |
| Font size | `w:sz` | ✅ | |
| Font color | `w:color` | ✅ | |
| Highlight | `w:highlight` | ✅ | |
| Subscript/superscript | `w:vertAlign` | ✅ | |
| All caps | `w:caps` | ✅ | |
| Small caps | `w:smallCaps` | ✅ | |
| Hidden | `w:vanish` | ✅ | |
| Emboss | `w:emboss` | ✅ | |
| Imprint | `w:imprint` | ✅ | |
| Outline | `w:outline` | ✅ | |
| Shadow | `w:shadow` | ✅ | |
| Character spacing | `w:spacing` | ❌ | |
| Text effects (w14) | `w14:textFill`, etc. | ❌ | |

### 2.6 Paragraph Formatting

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Alignment | `w:jc` | ✅ | |
| Indentation | `w:ind` | ✅ | |
| Spacing | `w:spacing` | ✅ | |
| Line spacing | `w:spacing/@w:line` | ✅ | |
| Tab stops | `w:tabs` | ✅ | |
| Borders | `w:pBdr` | ⚠️ | |
| Shading | `w:shd` | ⚠️ | |
| Keep together | `w:keepLines` | ✅ | |
| Keep with next | `w:keepNext` | ✅ | |
| Page break before | `w:pageBreakBefore` | ✅ | |
| Widow/orphan control | `w:widowControl` | ✅ | |
| Outline level | `w:outlineLvl` | ✅ | |
| Numbering | `w:numPr` | ⚠️ | Read only |

### 2.7 Styles

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Paragraph styles | `w:style[@w:type='paragraph']` | ✅ | |
| Character styles | `w:style[@w:type='character']` | ✅ | |
| Table styles | `w:style[@w:type='table']` | ✅ | |
| Numbering styles | `w:style[@w:type='numbering']` | ⚠️ | |
| Style hierarchy | `w:basedOn` | ✅ | |
| Next style | `w:next` | ✅ | |
| Latent styles | `w:latentStyles` | ✅ | |
| Default styles | | ✅ | |

### 2.8 Numbering

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Read numbering | `w:numbering` | ✅ | |
| Apply numbering | `w:numPr` | ✅ | |
| Create abstract num | `w:abstractNum` | ❌ | High priority |
| Create num instance | `w:num` | ⚠️ | |
| Level overrides | `w:lvlOverride` | ✅ | |
| Restart numbering | `w:startOverride` | ✅ | |

### 2.9 Images and Drawings

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Inline pictures | `wp:inline` | ✅ | |
| Floating pictures | `wp:anchor` | ⚠️ | Registered, minimal |
| Picture properties | `pic:pic` | ✅ | |
| Image sizing | | ✅ | |
| PNG | | ✅ | |
| JPEG | | ✅ | |
| TIFF | | ✅ | |
| BMP | | ✅ | |
| GIF | | ✅ | |
| SVG | | ❌ | |
| EMF/WMF | | ❌ | |
| Charts | `c:chart` | 🔍 | Detected only |
| SmartArt | `dgm:*` | 🔍 | Detected only |
| Shapes | `wps:wsp` | ❌ | |
| Text boxes | | ❌ | |
| WordArt | | ❌ | |

### 2.10 Document Parts

| Feature | Part Type | Status | Notes |
|---------|-----------|--------|-------|
| Main document | document.xml | ✅ | |
| Styles | styles.xml | ✅ | |
| Numbering | numbering.xml | ⚠️ | Read only |
| Settings | settings.xml | ⚠️ | Limited |
| Core properties | core.xml | ✅ | |
| Comments | comments.xml | ✅ | v1.2.0 |
| Headers | header*.xml | ✅ | |
| Footers | footer*.xml | ✅ | |
| Footnotes | footnotes.xml | ❌ | |
| Endnotes | endnotes.xml | ❌ | |
| Font table | fontTable.xml | ❌ | |
| Web settings | webSettings.xml | ❌ | |
| Theme | theme*.xml | ❌ | |
| Glossary | glossary.xml | ❌ | |
| Custom XML | customXml*.xml | ❌ | |

### 2.11 Advanced Features

| Feature | Element(s) | Status | Notes |
|---------|-----------|--------|-------|
| Track changes | `w:ins`, `w:del`, etc. | ❌ | |
| Math equations | `m:oMath` | ⚠️ | Preserved, not accessible |
| Bibliography | `b:Sources` | ❌ | |
| Captions | | ❌ | |
| Table of contents | Field-based | ❌ | |
| Index | Field-based | ❌ | |
| Mail merge | `w:mailMerge` | ❌ | |
| Document protection | `w:documentProtection` | ❌ | |
| Digital signatures | | ❌ | |

---

## 3. Architecture Assessment

### Strengths

1. **Clean layered architecture**: API → Proxy → OXML → lxml → OPC
2. **Declarative XML binding**: xmlchemy.py provides elegant element mapping
3. **Solid OPC layer**: Full package handling, relationships, content types
4. **Good test coverage**: Unit tests + BDD acceptance tests
5. **Type hints**: Modern Python with strict pyright checking

### Technical Debt

1. **Namespace map incomplete**: Missing w15, w16, wp14 namespaces
2. **Settings part minimal**: Only `evenAndOddHeaders` exposed
3. **Many registered but unused elements**: Placeholders in `__init__.py`
4. **No strict/transitional handling**: Always assumes transitional

### Extension Points

The architecture supports extension via:

1. **New OXML classes**: Add CT_* classes in `oxml/`
2. **Register elements**: Add to `oxml/__init__.py`
3. **New parts**: Add XmlPart subclasses in `parts/`
4. **Proxy objects**: Add user-facing classes wrapping OXML

---

## 4. Priority Recommendations

### Immediate (Tier 1)
1. **Content Controls** — Enables forms, templates, automation
2. **Fields** — Dynamic content, page numbers, cross-refs
3. **Numbering creation** — Programmatic list creation
4. **Footnotes/Endnotes** — Academic and professional documents

### Near-term (Tier 2)
5. **Bookmarks** — Named locations, cross-references
6. **Track Changes (read)** — Revision detection
7. **Floating images** — Positioned graphics
8. **Modern namespaces** — w15/w16 compatibility

### Long-term (Tier 3)
9. **Charts** — Embedded chart manipulation
10. **SmartArt** — Diagram editing
11. **Themes** — Color/font scheme access
12. **Math equations** — OMML access

---

## 5. Sources

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ISO/IEC 29500-1:2016](https://www.iso.org/standard/71691.html)
- [ISO/IEC 29500-2:2021](https://www.iso.org/standard/77818.html)
- [MS-DOCX Word Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/)
- [Office Open XML - Wikipedia](https://en.wikipedia.org/wiki/Office_Open_XML)
- [OOXML Format Family - Library of Congress](https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml)
