# Tier 3 Feature Specifications

This document provides detailed specifications for Tier 3 backlog items based on research.

---

## B-NS-01: Word 2013+ Namespace Support

### Overview

Microsoft Word uses namespace extensions to add new features in each version. Documents created in modern Word versions contain elements in namespaces that python-docx doesn't currently recognize. While these elements are preserved during round-trip, they cannot be accessed or created programmatically.

### Current State

From `src/docx/oxml/ns.py`, only `w14` (Word 2010) is currently defined:
```python
nsmap = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    # ... other namespaces
}
```

### Missing Namespaces

| Prefix | Namespace URI | Word Version | Features |
|--------|---------------|--------------|----------|
| w15 | `http://schemas.microsoft.com/office/word/2012/wordml` | 2013 | Comments extended, web extensions |
| w16 | `http://schemas.microsoft.com/office/word/2018/wordml` | 2016+ | Comments ID extended |
| w16se | `http://schemas.microsoft.com/office/word/2015/wordml/symex` | 2016 | Symbol extensions |
| w16cid | `http://schemas.microsoft.com/office/word/2016/wordml/cid` | 2016+ | Comment IDs |
| w16cex | `http://schemas.microsoft.com/office/word/2020/wordml/cex` | 2020/365 | Comments extended |
| w16sdtdh | `http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash` | 2021/365 | SDT data hash |
| wp14 | `http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing` | 2010 | Anchor positioning |
| a14 | `http://schemas.microsoft.com/office/drawing/2010/main` | 2010 | Drawing extensions |
| asvg | `http://schemas.microsoft.com/office/drawing/2016/SVG/main` | 2016 | SVG support |

### Implementation Plan

1. Add namespace declarations to `nsmap` in `ns.py`
2. These are read-only additions - just enables xpath queries with these prefixes
3. No new element classes needed initially (elements already preserved via lxml)

### Testing

- Create test document in Word 2019/365 with modern features
- Verify namespace queries work
- Verify round-trip preserves elements

---

## B-IMG-01: SVG Image Support

### Overview

SVG (Scalable Vector Graphics) support was added in Office 2016/365. SVG images provide resolution-independent graphics that scale without quality loss.

### How SVG is Stored in OOXML

SVG images are stored with a **dual-reference approach** for backward compatibility:

1. **Primary image (PNG fallback)**: A raster PNG in `/word/media/` for older Word versions
2. **SVG image**: The actual SVG in `/word/media/` for modern Word versions

### XML Structure

```xml
<w:drawing>
  <wp:inline>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:blipFill>
            <!-- PNG fallback reference -->
            <a:blip r:embed="rId4">
              <a:extLst>
                <!-- SVG extension with reference to actual SVG -->
                <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
                  <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
                                r:embed="rId5"/>
                </a:ext>
              </a:extLst>
            </a:blip>
          </pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

### Content Types

| Type | Content Type |
|------|-------------|
| SVG | `image/svg+xml` |
| PNG fallback | `image/png` |

### Relationship Type

SVG uses the standard image relationship:
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`

### Implementation Plan

#### Phase 1: Reading SVG Images (Scope for B-IMG-01)

1. Add `asvg` namespace to `ns.py`
2. Add `image/svg+xml` content type to OPC constants
3. Create `Svg` image header class in `src/docx/image/svg.py`
4. Add SVG signature detection to `SIGNATURES` in `image/__init__.py`
5. Add `svg_blip` property to `CT_Blip` to access SVG extension

#### Phase 2: Writing SVG Images (Future)

1. Create PNG fallback automatically when adding SVG
2. Generate dual-reference structure
3. Register both parts and relationships

### SVG Header Parsing

SVG files are XML. Key properties:
- **Width/Height**: From root `<svg>` element `width`/`height` attributes or `viewBox`
- **DPI**: Default to 96 dpi (SVG standard)
- **Content-Type**: `image/svg+xml`

### Testing

- Create document with SVG in Word
- Read SVG properties (width, height)
- Verify type detection works
- Verify round-trip preserves SVG

---

## B-STY-01: Theme Support

### Overview

Document themes control the visual appearance including colors, fonts, and effects. Every Word document has a theme part (`word/theme/theme1.xml`) that defines these settings.

### Theme Structure (DrawingML)

```xml
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <!-- Color Scheme: 12 theme colors -->
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>

    <!-- Font Scheme: Major and Minor fonts -->
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>

    <!-- Format Scheme: Fill, line, and effect styles -->
    <a:fmtScheme name="Office">
      <a:fillStyleLst>...</a:fillStyleLst>
      <a:lnStyleLst>...</a:lnStyleLst>
      <a:effectStyleLst>...</a:effectStyleLst>
      <a:bgFillStyleLst>...</a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
```

### Theme Colors

| Element | Description | Typical Use |
|---------|-------------|-------------|
| dk1 | Dark 1 | Main text color |
| lt1 | Light 1 | Background |
| dk2 | Dark 2 | Secondary text |
| lt2 | Light 2 | Secondary background |
| accent1-6 | Accent colors | Headings, charts, highlights |
| hlink | Hyperlink | Unvisited links |
| folHlink | Followed hyperlink | Visited links |

### Theme Fonts

| Element | Description | Typical Use |
|---------|-------------|-------------|
| majorFont | Headings font | Heading styles |
| minorFont | Body font | Normal text, paragraphs |

### Relationship Type

`http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme`

### Content Type

`application/vnd.openxmlformats-officedocument.theme+xml`

### Implementation Plan

#### Phase 1: Reading Theme (Scope for B-STY-01)

1. Create `ThemePart` class in `src/docx/parts/theme.py`
2. Create OXML classes for theme elements:
   - `CT_Theme` (root element)
   - `CT_ThemeElements`
   - `CT_ColorScheme` with 12 color properties
   - `CT_FontScheme` with major/minor fonts
3. Create proxy classes:
   - `Theme` - main API
   - `ThemeColors` - access to 12 theme colors
   - `ThemeFonts` - access to major/minor fonts
4. Add `theme` property to `Document` class

#### Phase 2: Theme Color Resolution (Future)

Resolve `w:themeColor` attributes in styles to actual RGB values:
```xml
<w:color w:val="4F81BD" w:themeColor="accent1"/>
```

### API Design

```python
# Access theme
theme = document.theme

# Theme name
print(theme.name)  # "Office"

# Color scheme
print(theme.colors.accent1)  # RGBColor(0x44, 0x72, 0xC4)
print(theme.colors.dark1)    # RGBColor(0x00, 0x00, 0x00)

# Font scheme
print(theme.fonts.major_latin)  # "Calibri Light"
print(theme.fonts.minor_latin)  # "Calibri"
```

### Testing

- Load document with theme
- Read theme colors, verify RGB values
- Read font names
- Verify round-trip preserves theme

---

## B-DRW-05: Floating Shape Creation

### Overview

Floating (anchored) images can be positioned anywhere on the page, unlike inline images which flow with text. Currently python-docx only supports adding inline images.

### OOXML Structure

Floating images use `wp:anchor` instead of `wp:inline`:

```xml
<w:drawing>
  <wp:anchor distT="0" distB="0" distL="114300" distR="114300"
             simplePos="0" relativeHeight="251658240"
             behindDoc="0" locked="0" layoutInCell="1"
             allowOverlap="1">
    <!-- Position -->
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="column">
      <wp:posOffset>914400</wp:posOffset>
    </wp:positionH>
    <wp:positionV relativeFrom="paragraph">
      <wp:posOffset>457200</wp:posOffset>
    </wp:positionV>

    <!-- Size -->
    <wp:extent cx="1828800" cy="1371600"/>

    <!-- Wrap style -->
    <wp:wrapSquare wrapText="bothSides"/>

    <!-- Content -->
    <a:graphic>
      <a:graphicData uri="...picture">
        <pic:pic>...</pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>
```

### Position Options

**Horizontal (relativeFrom):**
- `character`, `column`, `insideMargin`, `leftMargin`, `margin`, `outsideMargin`, `page`, `rightMargin`

**Vertical (relativeFrom):**
- `insideMargin`, `line`, `margin`, `outsideMargin`, `page`, `paragraph`, `topMargin`, `bottomMargin`

**Alignment (instead of offset):**
- Horizontal: `left`, `center`, `right`, `inside`, `outside`
- Vertical: `top`, `center`, `bottom`, `inside`, `outside`

### Wrap Styles

| Element | Description |
|---------|-------------|
| `wp:wrapNone` | No wrapping |
| `wp:wrapSquare` | Square wrapping |
| `wp:wrapTight` | Tight wrapping |
| `wp:wrapThrough` | Through wrapping |
| `wp:wrapTopAndBottom` | Top and bottom only |

### Proposed API

```python
# Add floating image with absolute position
shape = document.add_floating_picture(
    'image.png',
    width=Inches(2),
    height=Inches(1.5),
    left=Inches(1),      # from left margin
    top=Inches(2),       # from top of paragraph
    wrap='square'        # wrapping style
)

# Add floating image with alignment
shape = document.add_floating_picture(
    'image.png',
    width=Inches(2),
    horizontal='center',
    horizontal_relative_to='page',
    vertical='top',
    vertical_relative_to='margin',
    wrap='tight'
)

# Behind text
shape = document.add_floating_picture(
    'image.png',
    behind_text=True
)
```

### Implementation Plan

1. Create `CT_Anchor.new()` factory method
2. Add position elements (`positionH`, `positionV`, `simplePos`)
3. Add wrap elements
4. Create `add_floating_picture()` method on `Document`
5. Handle `relativeHeight` (z-order) management

### Acceptance Criteria

- [ ] Create anchored images with absolute position
- [ ] Set horizontal/vertical alignment
- [ ] Set wrap style (square, tight, none, etc.)
- [ ] Set behind_text property
- [ ] Proper z-order handling

### References

- [Floating Position](http://officeopenxml.com/drwPicFloating-position.php)
- [wp:anchor Specification](https://www.datypic.com/sc/ooxml/e-wp_anchor.html)
- [Anchor Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor)

---

## B-STY-02: Theme Modification

### Overview

Extend existing theme read support to allow modification of theme colors and fonts.

### Current State

Theme reading is implemented in `src/docx/theme.py`:
- `doc.theme.colors.accent1` - Read theme colors
- `doc.theme.fonts.major_latin` - Read font names

### Proposed API

```python
# Modify colors
doc.theme.colors.accent1 = RGBColor(0xFF, 0x00, 0x00)
doc.theme.colors.dark1 = RGBColor(0x00, 0x00, 0x00)

# Modify fonts
doc.theme.fonts.major_latin = 'Arial'
doc.theme.fonts.minor_latin = 'Times New Roman'

# Save changes
doc.save('modified.docx')
```

### Implementation Plan

1. Add setters to `ThemeColors` properties
2. Add setters to `ThemeFonts` properties
3. Create/modify OXML elements as needed
4. Handle system colors vs sRGB colors

### XML Modifications

**Color change:**
```xml
<!-- Before -->
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>

<!-- After -->
<a:accent1><a:srgbClr val="FF0000"/></a:accent1>
```

**Font change:**
```xml
<!-- Before -->
<a:latin typeface="Calibri Light"/>

<!-- After -->
<a:latin typeface="Arial"/>
```

### Acceptance Criteria

- [ ] Modify all 12 theme colors
- [ ] Modify major and minor fonts
- [ ] Changes persist after save
- [ ] Document displays correctly in Word

---

## B-CORE-02: Strict OOXML Conformance Detection

### Overview

OOXML has two conformance classes: Transitional (default) and Strict. They use different namespaces. python-docx assumes Transitional.

### Namespace Differences

| Type | Main Namespace |
|------|----------------|
| Transitional | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` |
| Strict | `http://purl.oclc.org/ooxml/wordprocessingml/main` |

### Detection Method

Check the namespace of the root element in `document.xml`:

```python
def detect_conformance(document_part):
    root = document_part._element
    ns = root.nsmap.get('w', '')
    if 'purl.oclc.org' in ns:
        return 'strict'
    return 'transitional'
```

### Word Version Detection

Based on `mc:Ignorable` attribute:
- No w14: Word 2007
- w14: Word 2010+
- w15: Word 2013+
- w16: Word 2016+
- w16cid, w16cex: Word 365/2019+
- w16sdtdh: Word 2021+

### Proposed API

```python
# Document conformance
print(doc.conformance)  # 'transitional' or 'strict'

# Minimum Word version required
print(doc.minimum_word_version)  # 'Word 2016' etc.

# Supported namespaces
print(doc.supported_namespaces)  # ['w14', 'w15', 'w16']
```

### Implementation Plan

1. Add conformance detection in `Document.__init__` or lazy property
2. Parse `mc:Ignorable` to detect Word version
3. Add warning for Strict documents
4. Consider namespace aliasing for Strict support

### Acceptance Criteria

- [ ] Detect Transitional vs Strict conformance
- [ ] Detect minimum Word version from namespaces
- [ ] Warning when opening Strict documents
- [ ] Basic Strict document reading (future: full support)

---

## References

### Specifications
- [ECMA-376: Office Open XML File Formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ISO 29500: Information technology — Document description and processing languages](https://www.iso.org/standard/71691.html)

### Documentation
- [Theme Part - OOXML](https://c-rex.net/projects/samples/ooxml/e1/Part1/OOXML_P1_Fundamentals_Theme_topic_ID0EUYNM.html)
- [a:theme element documentation](http://www.datypic.com/sc/ooxml/e-a_theme.html)
- [Theme Class - Open XML SDK](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme)
- [ThemeElements Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.themeelements)

### SVG Support
- [OpenXML insert SVG image](https://www.eidias.com/blog/2022/9/14/openxml-insert-svg-image-into-word-document)
- [SVG support in Office](https://office-watch.com/2016/svg-graphics-coming-to-office-at-long-last/)
- [Open XML SDK SVG Issue](https://github.com/OfficeDev/Open-XML-SDK/issues/204)
