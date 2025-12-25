# OOXML Version Compatibility Guide

**Last updated:** 2024-12-03 (T-0002 research)

This document covers OOXML versioning, Word version differences, and how to handle version compatibility in python-docx extensions.

---

## Executive Summary

OOXML has evolved significantly since Word 2007. Key challenges:

1. **Strict vs Transitional**: Two conformance classes with different schemas
2. **Namespace extensions**: w14, w15, w16, etc. add features per Word version
3. **Markup Compatibility (MCE)**: `mc:AlternateContent` wraps version-specific content
4. **python-docx gap**: Content inside `mc:AlternateContent` is invisible to `doc.paragraphs`

**Critical finding**: python-docx currently loses access to paragraphs/runs inside `mc:AlternateContent` blocks. This affects text boxes, modern vector graphics, and Word 2010+ specific features.

---

## OOXML Conformance Classes

### Transitional (Default)

- **What it is**: The backwards-compatible format, default in Word
- **Namespace**: `http://schemas.openxmlformats.org/wordprocessingml/2006/main`
- **Features**: Supports legacy VML graphics, deprecated elements
- **Used by**: All Word versions by default when saving

### Strict

- **What it is**: Pure OOXML without legacy elements
- **Namespace**: `http://purl.oclc.org/ooxml/wordprocessingml/main`
- **Features**: No VML, no deprecated elements
- **Used by**: Must be explicitly selected, rare in practice
- **Word support**: Fully supported only in Word 2013+

### Implications for python-docx

python-docx currently assumes Transitional format. The namespaces in `docx/oxml/ns.py` are all Transitional. Supporting Strict would require:

1. Detecting the conformance class
2. Alternative namespace mappings
3. Potentially different element handling

**Recommendation**: Focus on Transitional (95%+ of real documents). Add Strict detection with a warning.

---

## Word Version Namespace Extensions

Each Word version adds new namespaces for new features:

| Namespace | Prefix | Word Version | Example Features |
|-----------|--------|--------------|------------------|
| wordprocessingml/2006/main | w | 2007 | Core document |
| drawingml/2010/main | a14 | 2010 | Drawing extensions |
| wordprocessingml/2010/wordml | w14 | 2010 | Content controls, text effects |
| wordprocessingml/2012/wordml | w15 | 2013 | Charts, web extensions |
| wordprocessingml/2015/wordml | w16 | 2016 | Comments extended |
| wordprocessingExSe/2015/wordml | w16se | 2016 | Comments IDs |
| wordprocessingCid/2016/wordml | w16cid | 2016+ | Comment IDs (extended) |
| wordprocessingCex/2020/wordml | w16cex | 2020/365 | Comments extended |
| wordprocessingSdtDh/2020/wordml | w16sdtdh | 2021/365 | SDT date handling |
| drawingml/2010/wordprocessingDrawing | wp14 | 2010 | Anchor positioning |

### Current python-docx Support

From `src/docx/oxml/ns.py`:
```python
nsmap = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # ... plus mc, m, wne, wpc, wpg, wps
}
```

**Missing namespaces**: w14, w15, w16, w16se, w16cid, w16cex, w16sdtdh, wp14, a14

### Impact of Missing Namespaces

1. **Reading**: Elements in unknown namespaces are preserved but inaccessible
2. **Writing**: Cannot create elements requiring these namespaces
3. **mc:Ignorable**: Word lists supported namespaces; if we save with one unlisted, Word may fail to open

---

## Markup Compatibility and Extensibility (MCE)

MCE (ECMA-376 Part 3, ISO 29500-3) allows documents to contain version-specific content with fallbacks.

### Key Elements

| Element | Purpose |
|---------|---------|
| `mc:AlternateContent` | Container for version-specific alternatives |
| `mc:Choice` | Preferred content if `@Requires` namespaces are supported |
| `mc:Fallback` | Fallback content for older readers |

### Key Attributes

| Attribute | Purpose |
|-----------|---------|
| `mc:Ignorable` | Namespaces that can be ignored if unknown |
| `mc:ProcessContent` | Ignored elements whose children should still be processed |
| `mc:PreserveElements` | Elements to preserve even if in ignored namespace |
| `mc:PreserveAttributes` | Attributes to preserve even if in ignored namespace |

### Example: Text Box in Modern Word

```xml
<mc:AlternateContent>
  <mc:Choice Requires="wps">
    <w:drawing>
      <wp:anchor>
        <a:graphic>
          <a:graphicData>
            <wps:wsp>
              <wps:txbx>
                <w:txbxContent>
                  <w:p><w:r><w:t>Text here</w:t></w:r></w:p>
                </w:txbxContent>
              </wps:txbx>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </mc:Choice>
  <mc:Fallback>
    <w:pict>
      <v:shape>
        <v:textbox>
          <w:txbxContent>
            <w:p><w:r><w:t>Text here</w:t></w:r></w:p>
          </w:txbxContent>
        </v:textbox>
      </v:shape>
    </w:pict>
  </mc:Fallback>
</mc:AlternateContent>
```

### The python-docx Problem

**Issue #1389**: `doc.paragraphs` does not include content inside `mc:AlternateContent`.

Current behavior:
```python
# If paragraph is inside <mc:AlternateContent>:
for p in doc.paragraphs:
    print(p.text)  # The paragraph is MISSING from this list!
```

The document's body XPath `./w:p` only finds direct children, not paragraphs nested inside MCE blocks.

### How Other Libraries Handle MCE

#### docx4j (Java)
- Tries to use `mc:AlternateContent` when available
- Falls back to mc:Fallback content
- **On save**: Drops content it doesn't understand, effectively saving as Word 2007 compatible
- Recent versions (8.2.9+) preserve w16sdtdh namespace to avoid Word open failures

#### Open XML SDK (.NET)
- Has explicit `AlternateContent` class
- Provides `GetContentFromACBlock(block, FileFormatVersions)` method
- Can extract content for Office2007, Office2010, Office2013 targets
- Known issue: Root-level AlternateContent not properly loaded (Issue #1291)

#### Apache POI (Java)
- Uses ECMA-376 5th edition schemas
- Low-level XMLBeans access for content outside XWPF model
- Documentation acknowledges incomplete coverage

---

## Known Community Issues and Workarounds

### GitHub Issues Summary

| Issue | Topic | Status | Workaround |
|-------|-------|--------|------------|
| [#1389](https://github.com/python-openxml/python-docx/issues/1389) | mc:AlternateContent text invisible | Open | Navigate XML manually |
| [#155](https://github.com/python-openxml/python-docx/issues/155) | SDT/Content Controls | Open | Raw XML access |
| [#31](https://github.com/python-openxml/python-docx/issues/31) | Field codes | Open | Template with placeholders |
| [#1](https://github.com/python-openxml/python-docx/issues/1) | Footnotes | Open | PR #624, bayoo-docx fork |
| [#761](https://github.com/python-openxml/python-docx/issues/761) | SDT checkboxes | Open | None |

### Workarounds in Use

#### 1. Manual XML Navigation
```python
from lxml import etree

body = doc.element.body
xml_str = etree.tostring(body)
tree = etree.fromstring(xml_str)

# Find all paragraphs including those in mc:AlternateContent
ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
}
all_paragraphs = tree.xpath('.//w:p', namespaces=ns)
```

#### 2. TextBox Content Extraction
```python
XPATH_CHOICE = "./mc:Choice/w:drawing/wp:anchor/a:graphic/a:graphicData//wps:txbx/w:txbxContent"
XPATH_FALLBACK = "./mc:Fallback/w:pict//v:textbox/w:txbxContent"
```

#### 3. Template Approach for Fields
Create document in Word with fields, use python-docx to fill in surrounding text, use win32com or LibreOffice to update fields.

#### 4. bayoo-docx Fork
The [bayoo-docx](https://github.com/nicholasceliano/bayoo-docx) fork (based on PR #624) adds footnotes and extended comments support.

---

## Recommendations for WOTAN Development

### 1. Add Missing Namespaces
Add to `ns.py`:
```python
'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
'w16': 'http://schemas.microsoft.com/office/word/2018/wordml',
'w16se': 'http://schemas.microsoft.com/office/word/2015/wordmlExSe',
'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordmlCid',
'w16cex': 'http://schemas.microsoft.com/office/word/2020/wordmlCex',
'w16sdtdh': 'http://schemas.microsoft.com/office/word/2020/wordmlSdtDh',
'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
```

### 2. Handle mc:AlternateContent
Options:
- **Option A (Simple)**: Flatten mc:AlternateContent by selecting mc:Choice content
- **Option B (Full)**: Expose MCE structure, let caller choose version
- **Option C (Smart)**: Auto-select based on detected document version

Recommend starting with Option A for reading, preserving original on save.

### 3. Preserve mc:Ignorable on Save
When saving, ensure `mc:Ignorable` in document.xml lists all namespaces that appear in the document. Otherwise Word may fail to open.

### 4. Add Version Detection
```python
def detect_word_version(document):
    """Detect the minimum Word version that created this document."""
    ignorable = doc.element.get('{...}Ignorable', '')
    if 'w16sdtdh' in ignorable:
        return 'Word 365 (2021+)'
    elif 'w16cex' in ignorable:
        return 'Word 2020/365'
    elif 'w16' in ignorable:
        return 'Word 2016+'
    elif 'w15' in ignorable:
        return 'Word 2013+'
    elif 'w14' in ignorable:
        return 'Word 2010+'
    else:
        return 'Word 2007'
```

### 5. Test Matrix
Create test documents from each Word version to verify handling:
- Word 2007 (baseline)
- Word 2010 (w14)
- Word 2013 (w15)
- Word 2016/2019 (w16, w16se)
- Word 365/2021 (w16cid, w16cex, w16sdtdh)

---

## References

### Specifications
- [ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ISO 29500-3: Markup Compatibility and Extensibility](https://www.loc.gov/preservation/digital/formats/fdd/fdd000396.shtml)

### Articles
- [MCE Overview by Eric White](http://www.ericwhite.com/blog/markup-compatibility-and-extensibility/)
- [Apache OpenOffice MCE Wiki](https://wiki.openoffice.org/wiki/OOXML/Markup_Compatibility_and_Extensibility)
- [How to edit Microsoft Word documents in Python](https://www.rikvoorhaar.com/blog/python_docx) - TextBox workaround

### Library Documentation
- [docx4j Changelog](https://github.com/plutext/docx4j/blob/master/CHANGELOG.md) - Version compatibility updates
- [Open XML SDK AlternateContent](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.alternatecontent) - .NET handling
- [Apache POI XWPF](https://poi.apache.org/components/document/) - Java handling

### python-docx Issues
- [#1389: mc:AlternateContent content invisible](https://github.com/python-openxml/python-docx/issues/1389)
- [#155: SDT reading](https://github.com/python-openxml/python-docx/issues/155)
- [#1: Footnotes](https://github.com/python-openxml/python-docx/issues/1)
- [PR #624: Footnotes implementation](https://github.com/python-openxml/python-docx/pull/624)
