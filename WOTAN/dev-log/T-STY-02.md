# T-STY-02: Implement Theme Modification

| Field | Value |
|-------|-------|
| ID | T-STY-02 |
| Parent | B-STY-02 |
| State | DONE |
| Created | 2025-12-03 |
| Completed | 2025-12-03 |

## Objective

Enable modifying theme colors and fonts programmatically.

## Acceptance Criteria

- [x] Modify all 12 theme colors
- [x] Modify major and minor fonts
- [x] Changes persist after save
- [x] Document displays correctly in Word
- [x] All tests pass

## Context

**Specification:** `WOTAN/docs/tier3-specifications.md` - B-STY-02 section

**Current State:**
- Can read theme colors via `doc.theme.colors.accent1`
- Can read font names via `doc.theme.fonts.major_latin`
- Cannot modify theme colors or fonts

**XML Structure:**

Color change:
```xml
<!-- Before -->
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>

<!-- After -->
<a:accent1><a:srgbClr val="FF0000"/></a:accent1>
```

Font change:
```xml
<!-- Before -->
<a:latin typeface="Calibri Light"/>

<!-- After -->
<a:latin typeface="Arial"/>
```

## Implementation Notes

### Changes Made

1. **src/docx/oxml/theme.py**
   - Added `CT_Color.rgb_color` setter
   - Creates `<a:srgbClr>` element, replacing any existing color
   - Handles both sysClr and srgbClr child elements
   - Fixed type annotations to use `| None` for ZeroOrOne elements

2. **src/docx/theme.py**
   - Added `ThemeColors._set_color()` helper method
   - Added setters for all 12 color properties: dark1, light1, dark2, light2, accent1-6, hyperlink, followed_hyperlink
   - Accepts both `RGBColor` and hex string values (with or without '#')
   - Added `ThemeFonts._set_font()` helper method
   - Added setters for all 6 font properties: major_latin, major_east_asian, major_complex_script, minor_latin, minor_east_asian, minor_complex_script

### API

```python
from docx import Document
from docx.shared import RGBColor

doc = Document()

# Modify theme colors
doc.theme.colors.accent1 = RGBColor(0xFF, 0x00, 0x00)  # Red
doc.theme.colors.accent2 = '00FF00'  # Green via hex string
doc.theme.colors.hyperlink = '#0000FF'  # Blue via hex with #

# Modify theme fonts
doc.theme.fonts.major_latin = 'Arial'
doc.theme.fonts.minor_latin = 'Times New Roman'

doc.save('themed.docx')
```

## Obstacles

None encountered.

## Evidence

### Test Output

```
Testing Theme Modification:
==================================================
Theme name: Office Theme
Original accent1: 4F81BD
Original hyperlink: 0000FF
After set accent1 = RGBColor(255,0,0): FF0000
After set accent2 = "00FF00": 00FF00
After set hyperlink = "#0000FF": 0000FF

Original major_latin: Calibri
Original minor_latin: Cambria
After set major_latin = "Arial": Arial
After set minor_latin = "Times New Roman": Times New Roman

After save and reload:
accent1: FF0000
accent2: 00FF00
hyperlink: 0000FF
major_latin: Arial
minor_latin: Times New Roman

All tests passed!
```

### Test Results

- pytest: 1609 passed
- behave: 67 features, 650 scenarios passed
- pyright: 0 errors

## Outcome

DONE - Theme modification implemented with full round-trip support. All 12 theme colors and 6 theme fonts can be modified programmatically. Changes persist after save/reload.
