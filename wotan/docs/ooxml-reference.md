# OOXML Quick Reference

Quick reference for OOXML elements and structures relevant to python-docx development.

## Specification Documents

| Document | Description | URL |
|----------|-------------|-----|
| ECMA-376 | Office Open XML File Formats | https://ecma-international.org/publications-and-standards/standards/ecma-376/ |
| ISO/IEC 29500-1 | Fundamentals and Markup | https://www.iso.org/standard/71691.html |
| ISO/IEC 29500-2 | Open Packaging Conventions | https://www.iso.org/standard/77818.html |
| MS-DOCX | Word Extensions | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/ |
| MS-OE376 | Office Implementation | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/ |

## Namespaces

```xml
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
```

## Document Structure

### Package Structure (.docx)

```
docx-file.docx (ZIP archive)
├── [Content_Types].xml          # Content type definitions
├── _rels/
│   └── .rels                    # Package relationships
├── docProps/
│   ├── core.xml                 # Core properties (title, author)
│   └── app.xml                  # Application properties
└── word/
    ├── document.xml             # Main document content
    ├── styles.xml               # Style definitions
    ├── settings.xml             # Document settings
    ├── numbering.xml            # Numbering definitions
    ├── fontTable.xml            # Font table
    ├── webSettings.xml          # Web settings
    ├── footnotes.xml            # Footnotes
    ├── endnotes.xml             # Endnotes
    ├── comments.xml             # Comments
    ├── header1.xml              # Header parts
    ├── footer1.xml              # Footer parts
    ├── theme/
    │   └── theme1.xml           # Theme definitions
    ├── media/
    │   └── image1.png           # Embedded images
    └── _rels/
        └── document.xml.rels    # Document relationships
```

### Main Document

```xml
<w:document>
  <w:body>
    <w:p>...</w:p>           <!-- Paragraphs -->
    <w:tbl>...</w:tbl>       <!-- Tables -->
    <w:sdt>...</w:sdt>       <!-- Content controls -->
    <w:sectPr>...</w:sectPr> <!-- Final section properties -->
  </w:body>
</w:document>
```

## Common Elements

### Paragraph (`w:p`)

```xml
<w:p>
  <w:pPr>                    <!-- Paragraph properties -->
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="center"/>   <!-- Alignment -->
    <w:spacing w:before="240" w:after="120"/>
    <w:ind w:left="720"/>    <!-- Indentation -->
    <w:numPr>                <!-- Numbering -->
      <w:ilvl w:val="0"/>
      <w:numId w:val="1"/>
    </w:numPr>
  </w:pPr>
  <w:r>...</w:r>             <!-- Runs -->
  <w:hyperlink>...</w:hyperlink>
  <w:bookmarkStart/><w:bookmarkEnd/>
</w:p>
```

### Run (`w:r`)

```xml
<w:r>
  <w:rPr>                    <!-- Run properties -->
    <w:b/>                   <!-- Bold -->
    <w:i/>                   <!-- Italic -->
    <w:u w:val="single"/>    <!-- Underline -->
    <w:color w:val="FF0000"/>
    <w:sz w:val="24"/>       <!-- Font size (half-points) -->
    <w:rFonts w:ascii="Arial"/>
  </w:rPr>
  <w:t>Text content</w:t>    <!-- Text -->
  <w:br/>                    <!-- Break -->
  <w:drawing>...</w:drawing> <!-- Inline image -->
</w:r>
```

### Table (`w:tbl`)

```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="5000" w:type="pct"/>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="2500"/>
    <w:gridCol w:w="2500"/>
  </w:tblGrid>
  <w:tr>
    <w:tc>
      <w:tcPr>
        <w:tcW w:w="2500"/>
        <w:vMerge w:val="restart"/>  <!-- Vertical merge start -->
        <w:gridSpan w:val="2"/>      <!-- Horizontal merge -->
      </w:tcPr>
      <w:p>...</w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

### Content Control (`w:sdt`)

```xml
<w:sdt>
  <w:sdtPr>
    <w:tag w:val="fieldName"/>
    <w:alias w:val="Field Label"/>
    <w:id w:val="123456"/>
    <w:placeholder>
      <w:docPart w:val="DefaultPlaceholder"/>
    </w:placeholder>
    <!-- Type-specific elements -->
    <w:text/>                <!-- Plain text -->
    <w:date/>                <!-- Date picker -->
    <w:dropDownList/>        <!-- Dropdown -->
    <w:comboBox/>            <!-- Combo box -->
    <w14:checkbox/>          <!-- Checkbox (w14) -->
  </w:sdtPr>
  <w:sdtContent>
    <w:p>...</w:p>           <!-- Block content -->
    <!-- or -->
    <w:r>...</w:r>           <!-- Inline content -->
  </w:sdtContent>
</w:sdt>
```

### Fields

#### Simple Field

```xml
<w:fldSimple w:instr=" PAGE ">
  <w:r>
    <w:t>1</w:t>
  </w:r>
</w:fldSimple>
```

#### Complex Field

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> TOC \o "1-3" </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Table of Contents</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

### Bookmarks

```xml
<w:bookmarkStart w:id="0" w:name="myBookmark"/>
<w:r><w:t>Bookmarked text</w:t></w:r>
<w:bookmarkEnd w:id="0"/>
```

### Comments

```xml
<!-- In document -->
<w:commentRangeStart w:id="0"/>
<w:r><w:t>Commented text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r>
  <w:commentReference w:id="0"/>
</w:r>

<!-- In comments.xml -->
<w:comment w:id="0" w:author="Author" w:date="2024-01-01T00:00:00Z">
  <w:p>
    <w:r><w:t>Comment text</w:t></w:r>
  </w:p>
</w:comment>
```

### Footnotes/Endnotes

```xml
<!-- In document -->
<w:r>
  <w:footnoteReference w:id="1"/>
</w:r>

<!-- In footnotes.xml -->
<w:footnote w:id="1">
  <w:p>
    <w:r><w:t>Footnote text</w:t></w:r>
  </w:p>
</w:footnote>
```

### Drawing (Inline Image)

```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="1000000" cy="500000"/>  <!-- EMUs -->
    <wp:docPr id="1" name="Picture 1"/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="0" name="image.png"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="rId4"/>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="1000000" cy="500000"/>
            </a:xfrm>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

### Track Changes

```xml
<!-- Insertion -->
<w:ins w:id="0" w:author="Author" w:date="2024-01-01T00:00:00Z">
  <w:r><w:t>Inserted text</w:t></w:r>
</w:ins>

<!-- Deletion -->
<w:del w:id="1" w:author="Author" w:date="2024-01-01T00:00:00Z">
  <w:r><w:delText>Deleted text</w:delText></w:r>
</w:del>

<!-- Formatting change -->
<w:r>
  <w:rPr>
    <w:b/>
    <w:rPrChange w:id="2" w:author="Author">
      <w:rPr/>  <!-- Original formatting -->
    </w:rPrChange>
  </w:rPr>
  <w:t>Text with changed formatting</w:t>
</w:r>
```

## Measurements

| Unit | Description | Conversion |
|------|-------------|------------|
| EMU | English Metric Unit | 914400 EMU = 1 inch |
| Twip | 1/20 point | 1440 twips = 1 inch |
| Half-point | Font size unit | 2 half-points = 1 point |
| Fiftieths of percent | Width unit | 5000 = 100% |

## Common Field Codes

| Field | Description |
|-------|-------------|
| `PAGE` | Current page number |
| `NUMPAGES` | Total pages |
| `DATE` | Current date |
| `TIME` | Current time |
| `AUTHOR` | Document author |
| `TITLE` | Document title |
| `TOC` | Table of contents |
| `REF` | Cross-reference |
| `HYPERLINK` | Hyperlink |
| `MERGEFIELD` | Mail merge field |
| `IF` | Conditional |
| `SEQ` | Sequence number |
