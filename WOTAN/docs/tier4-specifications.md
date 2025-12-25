# Tier 4 Feature Specifications

Research and specifications for the `[NEEDS-SPEC]` backlog items. These represent complex OOXML features that require significant investigation before implementation.

## B-MATH-01: Math Equations (OMML)

### Overview

Office Math Markup Language (OMML) is the XML format for mathematical expressions in OOXML documents. Math equations appear as `<m:oMath>` or `<m:oMathPara>` elements within document content.

### Namespace

```
xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
```

### Key Elements

| Element | Description |
|---------|-------------|
| `m:oMath` | Inline math zone (within a paragraph) |
| `m:oMathPara` | Math paragraph (block-level, centered) |
| `m:f` | Fraction (numerator/denominator) |
| `m:rad` | Radical (square root, nth root) |
| `m:nary` | N-ary operator (summation, product, integral) |
| `m:sSub` | Subscript |
| `m:sSup` | Superscript |
| `m:sSubSup` | Sub-superscript (both) |
| `m:sPre` | Pre-sub-superscript |
| `m:m` | Matrix |
| `m:d` | Delimiter (parentheses, brackets) |
| `m:acc` | Accent (hat, bar, dot over character) |
| `m:bar` | Bar (overbar, underbar) |
| `m:box` | Box (invisible grouping) |
| `m:borderBox` | Border box (visible border) |
| `m:eqArr` | Equation array (aligned equations) |
| `m:func` | Function application (sin, cos, etc.) |
| `m:groupChr` | Group character (brace above/below) |
| `m:limLow` | Lower limit |
| `m:limUpp` | Upper limit |
| `m:phant` | Phantom (invisible placeholder) |
| `m:r` | Math run (text within math) |

### Document Structure

Math equations can appear:
1. Inline within paragraphs (`<w:p>` → `<m:oMath>`)
2. As block-level math paragraphs (`<w:p>` → `<m:oMathPara>` → `<m:oMath>`)

### Example XML

```xml
<!-- Fraction: a/b -->
<m:oMath>
  <m:f>
    <m:num><m:r><m:t>a</m:t></m:r></m:num>
    <m:den><m:r><m:t>b</m:t></m:r></m:den>
  </m:f>
</m:oMath>

<!-- Square root of x -->
<m:oMath>
  <m:rad>
    <m:radPr><m:degHide m:val="1"/></m:radPr>
    <m:deg/>
    <m:e><m:r><m:t>x</m:t></m:r></m:e>
  </m:rad>
</m:oMath>

<!-- Summation: Σ(i=1 to n) of x_i -->
<m:oMath>
  <m:nary>
    <m:naryPr><m:chr m:val="∑"/></m:naryPr>
    <m:sub><m:r><m:t>i=1</m:t></m:r></m:sub>
    <m:sup><m:r><m:t>n</m:t></m:r></m:sup>
    <m:e>
      <m:sSub>
        <m:e><m:r><m:t>x</m:t></m:r></m:e>
        <m:sub><m:r><m:t>i</m:t></m:r></m:sub>
      </m:sSub>
    </m:e>
  </m:nary>
</m:oMath>
```

### Implementation Strategy

**Phase 1: Read Support (Recommended First)**
- Parse `m:oMath` and `m:oMathPara` elements
- Expose via `Document.math_equations` collection
- Provide `text` property with plain-text approximation
- Provide `xml` property for raw XML access

**Phase 2: Create Support**
- Builder API for common math structures
- LaTeX-to-OMML conversion (optional, complex)

### Complexity Assessment

**High complexity** - OMML has deep nesting and many element types. Full implementation requires handling 20+ element types with complex properties. Consider limiting initial scope to read-only access with text extraction.

### References

- ECMA-376 Part 1, Section 22.1 (Math)
- [Office Open XML Math Overview](http://www.officeopenxml.com/)

---

## B-CHART-01: Chart Support

### Overview

Charts in OOXML documents are stored as DrawingML chart objects within embedded SpreadsheetML packages. A chart consists of a drawing anchor, a chart reference, and the chart data itself.

### Namespaces

```
xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
```

### Document Structure

Charts are stored in separate parts:
```
/word/document.xml          - Contains drawing anchor with chart reference
/word/charts/chart1.xml     - Chart definition
/word/embeddings/sheet1.xlsx - Embedded Excel data (SpreadsheetML)
```

### Key Elements

**In document.xml:**
```xml
<w:drawing>
  <wp:inline>  <!-- or wp:anchor for floating -->
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart r:id="rId5"/>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

**In chart1.xml:**

| Element | Description |
|---------|-------------|
| `c:chartSpace` | Root element for chart |
| `c:chart` | Chart definition container |
| `c:plotArea` | The plotting area |
| `c:barChart` | Bar/column chart |
| `c:lineChart` | Line chart |
| `c:pieChart` | Pie chart |
| `c:areaChart` | Area chart |
| `c:scatterChart` | Scatter/XY chart |
| `c:doughnutChart` | Doughnut chart |
| `c:radarChart` | Radar chart |
| `c:ser` | Data series |
| `c:cat` | Category axis data |
| `c:val` | Value axis data |
| `c:legend` | Chart legend |
| `c:title` | Chart title |

### Example Chart XML

```xml
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:title>
      <c:tx><c:rich>...</c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$5</c:f>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$5</c:f>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:barChart>
      <c:catAx>...</c:catAx>
      <c:valAx>...</c:valAx>
    </c:plotArea>
    <c:legend>...</c:legend>
  </c:chart>
</c:chartSpace>
```

### Data Storage

Chart data is stored in an embedded Excel file (`/word/embeddings/`). The chart references cells via formulas like `Sheet1!$A$2:$A$5`.

Options for data access:
1. Parse the embedded SpreadsheetML directly
2. Use `openpyxl` or similar library to read the embedded xlsx
3. Cache data values in the chart XML itself (some charts do this)

### Implementation Strategy

**Phase 1: Detection and Basic Read**
- Detect charts via `Document.charts` property
- Expose chart type, title, legend presence
- Count data series

**Phase 2: Data Access**
- Extract cached data values from chart XML
- Parse embedded SpreadsheetML for source data

**Phase 3: Create/Modify (Complex)**
- Create charts with embedded data
- Modify chart properties and data

### Complexity Assessment

**Very high complexity** - Charts involve multiple interconnected parts (drawing, chart, embedded spreadsheet). The data model is complex with many chart types and options. Consider partnering with a charting library or limiting to read-only detection initially.

### References

- ECMA-376 Part 1, Section 21.2 (DrawingML Charts)
- [Office Open XML DrawingML Overview](http://www.officeopenxml.com/drwOverview.php)

---

## B-SMART-01: SmartArt Support

### Overview

SmartArt graphics are complex diagram objects composed of multiple parts. A single SmartArt graphic can have up to 5 associated XML parts that define its layout, data, styling, and optional custom drawing.

### Namespaces

```
xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
```

### Document Structure

SmartArt is referenced via a `dgm:relIds` element that points to multiple parts:

```
/word/document.xml              - Drawing anchor with diagram reference
/word/diagrams/data1.xml        - Data model (required)
/word/diagrams/layout1.xml      - Layout definition (required)
/word/diagrams/colors1.xml      - Color scheme (optional)
/word/diagrams/quickStyle1.xml  - Style definition (optional)
/word/diagrams/drawing1.xml     - Custom drawing (optional, for modified shapes)
```

### Key Elements

**In document.xml:**
```xml
<w:drawing>
  <wp:inline>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
        <dgm:relIds
          r:dm="rId4"   <!-- data -->
          r:lo="rId5"   <!-- layout -->
          r:qs="rId6"   <!-- quickStyle -->
          r:cs="rId7"/> <!-- colors -->
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

**In data1.xml (Data Model):**

| Element | Description |
|---------|-------------|
| `dgm:dataModel` | Root element |
| `dgm:ptLst` | Point list (data nodes) |
| `dgm:pt` | Individual data point/node |
| `dgm:t` | Text content of a node |
| `dgm:cxnLst` | Connection list |
| `dgm:cxn` | Connection between nodes |

**In layout1.xml (Layout Definition):**

| Element | Description |
|---------|-------------|
| `dgm:layoutDef` | Layout definition root |
| `dgm:layoutNode` | Layout node (building block) |
| `dgm:alg` | Algorithm (linear, cycle, hierarchy, etc.) |
| `dgm:shape` | Shape to render |
| `dgm:constrLst` | Constraint list |
| `dgm:forEach` | Iterator over data |
| `dgm:presOf` | Presentation binding to data |

### Example Data Model XML

```xml
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc">
      <dgm:prSet/>
    </dgm:pt>
    <dgm:pt modelId="1">
      <dgm:prSet/>
      <dgm:t>
        <a:bodyPr/>
        <a:p><a:r><a:t>First Item</a:t></a:r></a:p>
      </dgm:t>
    </dgm:pt>
    <dgm:pt modelId="2">
      <dgm:prSet/>
      <dgm:t>
        <a:bodyPr/>
        <a:p><a:r><a:t>Second Item</a:t></a:r></a:p>
      </dgm:t>
    </dgm:pt>
  </dgm:ptLst>
  <dgm:cxnLst>
    <dgm:cxn modelId="3" srcId="0" destId="1" type="parOf"/>
    <dgm:cxn modelId="4" srcId="0" destId="2" type="parOf"/>
  </dgm:cxnLst>
</dgm:dataModel>
```

### Layout Algorithms

SmartArt supports various layout algorithms:
- **linear** - Sequential arrangement
- **cycle** - Circular arrangement
- **hierarchy** - Tree structure
- **pyramid** - Pyramid shape
- **composite** - Combined layouts

### Implementation Strategy

**Phase 1: Detection and Text Extraction**
- Detect SmartArt via `Document.smartart` or `Document.diagrams`
- Extract text content from data model
- Identify diagram type/layout

**Phase 2: Data Model Access**
- Expose nodes and connections
- Read node text and properties
- Navigate hierarchy

**Phase 3: Modification (Complex)**
- Add/remove nodes
- Modify text and connections
- Layout recalculation not recommended (let Word handle)

### Complexity Assessment

**Very high complexity** - SmartArt involves 4-5 interconnected parts with complex layout algorithms. The layout definition language is essentially a domain-specific programming language. Recommend limiting to text extraction and data model read/write, leaving layout to Word.

### References

- ECMA-376 Part 1, Section 21.4 (DrawingML Diagrams)
- [Microsoft: Create Custom SmartArt Graphics](https://learn.microsoft.com/en-us/archive/msdn-magazine/2007/february/create-custom-smartart-graphics-for-use-in-the-2007-office-system)
- [Office Open XML DrawingML Overview](http://www.officeopenxml.com/drwOverview.php)

---

## B-XML-01: Custom XML Support

### Overview

Custom XML parts allow embedding arbitrary XML data within OOXML documents. This data can be bound to content controls for data-driven documents and templates.

### Namespaces

```
xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"
xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```

### Document Structure

Custom XML is stored in dedicated parts:
```
/customXml/item1.xml           - The custom XML data
/customXml/itemProps1.xml      - Properties (GUID, schema refs)
/customXml/_rels/item1.xml.rels - Relationship to properties
```

### Key Elements

**In itemProps1.xml:**

| Element | Description |
|---------|-------------|
| `ds:datastoreItem` | Root element for custom XML properties |
| `ds:itemID` | Unique GUID identifying this custom XML part |
| `ds:schemaRefs` | List of schema references |
| `ds:schemaRef` | Individual schema URI reference |

**In document.xml (Inline Custom XML):**

| Element | Description |
|---------|-------------|
| `w:customXml` | Inline custom XML wrapper |
| `w:customXmlPr` | Properties for inline custom XML |
| `w:element` | Element name |
| `w:uri` | Namespace URI |

**Content Control Data Binding:**

| Element | Description |
|---------|-------------|
| `w:dataBinding` | Binds content control to custom XML |
| `w:storeItemID` | GUID of the custom XML part |
| `w:xpath` | XPath expression to bind to |
| `w:prefixMappings` | Namespace prefix declarations |

### Example: Custom XML Part

**item1.xml:**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<customer xmlns="http://example.com/customer">
  <name>ACME Corporation</name>
  <contact>John Doe</contact>
  <email>john@acme.com</email>
</customer>
```

**itemProps1.xml:**
```xml
<ds:datastoreItem
    ds:itemID="{C66AC4F1-97D3-4D10-8D07-51E6CC79AD91}"
    xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs>
    <ds:schemaRef ds:uri="http://example.com/customer"/>
  </ds:schemaRefs>
</ds:datastoreItem>
```

### Example: Content Control Binding

```xml
<w:sdt>
  <w:sdtPr>
    <w:tag w:val="customer_name"/>
    <w:dataBinding
      w:storeItemID="{C66AC4F1-97D3-4D10-8D07-51E6CC79AD91}"
      w:xpath="/customer/name"
      w:prefixMappings="xmlns:ns0='http://example.com/customer'"/>
  </w:sdtPr>
  <w:sdtContent>
    <w:p><w:r><w:t>ACME Corporation</w:t></w:r></w:p>
  </w:sdtContent>
</w:sdt>
```

### Use Cases

1. **Document Templates**: Bind content controls to custom XML for mail-merge-like functionality
2. **Metadata Storage**: Store application-specific data within the document
3. **Data Exchange**: Export/import structured data to/from documents
4. **Document Generation**: Populate templates from external data sources

### Implementation Strategy

**Phase 1: Read Custom XML Parts**
- Access custom XML parts via `Document.custom_xml_parts`
- Parse and expose XML data
- List schema references

**Phase 2: Create/Modify Custom XML**
- Add new custom XML parts
- Modify existing custom XML data
- Set schema references

**Phase 3: Data Binding Integration**
- Connect to existing SDT data binding support
- Expose bound custom XML via content controls
- Update content when custom XML changes

### Complexity Assessment

**Medium complexity** - Custom XML parts are straightforward to read/write. The complexity lies in the data binding integration with content controls and maintaining consistency between bound values. Consider implementing in conjunction with enhanced SDT support.

### References

- ECMA-376 Part 1, Section 15.2.4 (Custom XML Data Storage)
- [Microsoft: Custom XML Parts Overview](https://learn.microsoft.com/en-us/visualstudio/vsto/custom-xml-parts-overview)
- [SAP: OpenXML Custom XML Part Mapping](https://blogs.sap.com/2017/04/24/openxml-in-word-processing-custom-xml-part-mapping-flat-data/)
- [TextControl: Custom XML Parts in Word](https://www.textcontrol.com/blog/2024/07/23/read-and-write-custom-xml-parts-in-ms-word-office-open-xml-docx-files-using-net-csharp/)

---

## Implementation Priority Recommendations

Based on complexity and utility:

| Feature | Complexity | Utility | Recommendation |
|---------|------------|---------|----------------|
| **Custom XML** | Medium | High | Implement first - enables data-driven templates |
| **Math Equations** | High | Medium | Read-only with text extraction |
| **Charts** | Very High | Medium | Detection + metadata only initially |
| **SmartArt** | Very High | Low-Medium | Detection + text extraction only |

### Suggested Order

1. **B-XML-01**: Custom XML Support (medium effort, high value)
2. **B-MATH-01**: Math Equations - read-only (high effort, medium value)
3. **B-CHART-01**: Charts - detection only (medium effort for detection)
4. **B-SMART-01**: SmartArt - detection only (medium effort for detection)

Full chart and SmartArt support would require significant investment and may be better served by integration with specialized libraries.
