"""Document generators for round-trip testing.

These generators create documents with various feature combinations to test
structural equivalence through read-save-read cycles.
"""

from pathlib import Path
from typing import Optional

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor

# Test image path (a small image from the repo)
TEST_IMAGE = Path(__file__).parent.parent / "test_files" / "having-images.docx"


def create_basic_document() -> Document:
    """Create a basic document with paragraphs and simple formatting."""
    doc = Document()

    # Title
    title = doc.add_heading("Test Document - Basic Features", 0)

    # Regular paragraphs
    doc.add_paragraph("This is a regular paragraph with plain text.")

    # Paragraph with formatting
    p = doc.add_paragraph()
    p.add_run("This paragraph has ")
    bold_run = p.add_run("bold")
    bold_run.bold = True
    p.add_run(", ")
    italic_run = p.add_run("italic")
    italic_run.italic = True
    p.add_run(", and ")
    both_run = p.add_run("both")
    both_run.bold = True
    both_run.italic = True
    p.add_run(" formatting.")

    # Paragraph with alignment
    centered = doc.add_paragraph("This paragraph is centered.")
    centered.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    right_aligned = doc.add_paragraph("This paragraph is right-aligned.")
    right_aligned.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    return doc


def create_table_document() -> Document:
    """Create a document with various table configurations."""
    doc = Document()

    doc.add_heading("Test Document - Tables", 0)

    # Simple table
    doc.add_paragraph("Simple 3x3 table:")
    table1 = doc.add_table(rows=3, cols=3)
    for i, row in enumerate(table1.rows):
        for j, cell in enumerate(row.cells):
            cell.text = f"R{i+1}C{j+1}"

    doc.add_paragraph()

    # Table with header row
    doc.add_paragraph("Table with header:")
    table2 = doc.add_table(rows=4, cols=3)
    table2.style = "Table Grid"
    headers = ["Name", "Age", "City"]
    for j, header in enumerate(headers):
        table2.rows[0].cells[j].text = header

    data = [
        ("Alice", "30", "New York"),
        ("Bob", "25", "London"),
        ("Carol", "35", "Paris"),
    ]
    for i, (name, age, city) in enumerate(data):
        table2.rows[i + 1].cells[0].text = name
        table2.rows[i + 1].cells[1].text = age
        table2.rows[i + 1].cells[2].text = city

    return doc


def create_nested_table_document() -> Document:
    """Create a document with nested tables."""
    doc = Document()

    doc.add_heading("Test Document - Nested Tables", 0)

    # Outer table
    outer = doc.add_table(rows=2, cols=2)
    outer.style = "Table Grid"

    outer.rows[0].cells[0].text = "Simple cell"
    outer.rows[0].cells[1].text = "Another simple cell"
    outer.rows[1].cells[0].text = "Cell with text"

    # Nested table in bottom-right cell
    nested_cell = outer.rows[1].cells[1]
    nested_cell.paragraphs[0].text = "Cell with nested table:"
    inner = nested_cell.add_table(rows=2, cols=2)
    for i, row in enumerate(inner.rows):
        for j, cell in enumerate(row.cells):
            cell.text = f"Inner {i},{j}"

    return doc


def create_styled_document() -> Document:
    """Create a document with various styles applied."""
    doc = Document()

    doc.add_heading("Test Document - Styles", 0)

    # Different heading levels
    doc.add_heading("Heading Level 1", 1)
    doc.add_paragraph("Content under heading 1.")

    doc.add_heading("Heading Level 2", 2)
    doc.add_paragraph("Content under heading 2.")

    doc.add_heading("Heading Level 3", 3)
    doc.add_paragraph("Content under heading 3.")

    # List styles
    doc.add_paragraph("Bullet item 1", style="List Bullet")
    doc.add_paragraph("Bullet item 2", style="List Bullet")
    doc.add_paragraph("Bullet item 3", style="List Bullet")

    doc.add_paragraph("Number item 1", style="List Number")
    doc.add_paragraph("Number item 2", style="List Number")
    doc.add_paragraph("Number item 3", style="List Number")

    # Quote style
    doc.add_paragraph(
        "This is a quote block with a longer text to demonstrate the quote style.",
        style="Quote",
    )

    return doc


def create_multi_section_document() -> Document:
    """Create a document with multiple sections and different page layouts."""
    doc = Document()

    # First section (portrait)
    doc.add_heading("Section 1 - Portrait", 0)
    doc.add_paragraph("This section uses portrait orientation.")
    doc.add_paragraph("It has default margins and layout.")

    # Add section break and switch to landscape
    new_section = doc.add_section()
    new_section.orientation = WD_ORIENT.LANDSCAPE
    # Swap width and height for landscape
    new_section.page_width, new_section.page_height = (
        new_section.page_height,
        new_section.page_width,
    )

    doc.add_heading("Section 2 - Landscape", 1)
    doc.add_paragraph("This section uses landscape orientation.")
    doc.add_paragraph("Good for wide tables or charts.")

    # Add another section (back to portrait)
    new_section2 = doc.add_section()
    new_section2.orientation = WD_ORIENT.PORTRAIT
    new_section2.page_width, new_section2.page_height = (
        new_section2.page_height,
        new_section2.page_width,
    )

    doc.add_heading("Section 3 - Portrait Again", 1)
    doc.add_paragraph("Back to portrait orientation.")

    return doc


def create_header_footer_document() -> Document:
    """Create a document with headers and footers."""
    doc = Document()

    # Get the default section
    section = doc.sections[0]

    # Add header
    header = section.header
    header_para = header.paragraphs[0]
    header_para.text = "Document Header - Test Document"
    header_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Add footer
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.text = "Page Footer - Confidential"
    footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Body content
    doc.add_heading("Document with Header and Footer", 0)
    doc.add_paragraph("This document has a header and footer defined.")
    doc.add_paragraph("The header should appear at the top of each page.")
    doc.add_paragraph("The footer should appear at the bottom of each page.")

    return doc


def create_font_formatting_document() -> Document:
    """Create a document with various font formatting options."""
    doc = Document()

    doc.add_heading("Test Document - Font Formatting", 0)

    # Font sizes
    p1 = doc.add_paragraph()
    run_small = p1.add_run("Small text (8pt) ")
    run_small.font.size = Pt(8)
    run_normal = p1.add_run("Normal text (11pt) ")
    run_normal.font.size = Pt(11)
    run_large = p1.add_run("Large text (18pt)")
    run_large.font.size = Pt(18)

    # Font colors
    p2 = doc.add_paragraph()
    run_red = p2.add_run("Red text ")
    run_red.font.color.rgb = RGBColor(255, 0, 0)
    run_green = p2.add_run("Green text ")
    run_green.font.color.rgb = RGBColor(0, 128, 0)
    run_blue = p2.add_run("Blue text")
    run_blue.font.color.rgb = RGBColor(0, 0, 255)

    # Font styles
    p3 = doc.add_paragraph()
    run_strike = p3.add_run("Strikethrough ")
    run_strike.font.strike = True
    run_sub = p3.add_run("Subscript")
    run_sub.font.subscript = True
    p3.add_run(" Normal ")
    run_sup = p3.add_run("Superscript")
    run_sup.font.superscript = True

    # Underline
    p4 = doc.add_paragraph()
    run_underline = p4.add_run("Underlined text")
    run_underline.font.underline = True

    return doc


def create_mixed_content_document() -> Document:
    """Create a document with mixed content types."""
    doc = Document()

    doc.add_heading("Test Document - Mixed Content", 0)

    # Paragraph
    doc.add_paragraph("This document contains various content types mixed together.")

    # Table
    doc.add_paragraph("Here's a table:")
    table = doc.add_table(rows=2, cols=3)
    table.style = "Table Grid"
    for i, row in enumerate(table.rows):
        for j, cell in enumerate(row.cells):
            cell.text = f"Cell {i},{j}"

    # More paragraphs
    doc.add_paragraph("After the table, more content follows.")

    # Heading
    doc.add_heading("Sub-section", 2)

    # Bullet list
    doc.add_paragraph("Item A", style="List Bullet")
    doc.add_paragraph("Item B", style="List Bullet")

    # Another table
    doc.add_paragraph("Another table with different content:")
    table2 = doc.add_table(rows=3, cols=2)
    table2.rows[0].cells[0].text = "Header 1"
    table2.rows[0].cells[1].text = "Header 2"
    table2.rows[1].cells[0].text = "Data A"
    table2.rows[1].cells[1].text = "Data B"
    table2.rows[2].cells[0].text = "Data C"
    table2.rows[2].cells[1].text = "Data D"

    # Closing paragraph
    doc.add_paragraph("End of mixed content document.")

    return doc


def create_unicode_document() -> Document:
    """Create a document with various Unicode characters."""
    doc = Document()

    doc.add_heading("Test Document - Unicode Content", 0)

    # Various languages
    doc.add_paragraph("English: Hello, World!")
    doc.add_paragraph("Swedish: Hej världen! Åäö ÅÄÖ")
    doc.add_paragraph("German: Grüß Gott! ß ü ö ä")
    doc.add_paragraph("French: Bonjour le monde! é è ê ë")
    doc.add_paragraph("Spanish: ¡Hola mundo! ñ ¿")
    doc.add_paragraph("Russian: Привет мир!")
    doc.add_paragraph("Chinese: 你好世界!")
    doc.add_paragraph("Japanese: こんにちは世界!")
    doc.add_paragraph("Arabic: مرحبا بالعالم!")
    doc.add_paragraph("Emoji: 🎉 🚀 💡 ✅ ❌")

    # Special characters
    doc.add_paragraph("Symbols: © ® ™ § ¶ † ‡ • … – —")
    doc.add_paragraph("Math: ∑ ∏ √ ∞ ≈ ≠ ≤ ≥ ± × ÷")
    doc.add_paragraph("Arrows: → ← ↑ ↓ ↔ ⇒ ⇐")

    return doc


def create_long_document() -> Document:
    """Create a longer document with repeated sections."""
    doc = Document()

    doc.add_heading("Test Document - Long Document", 0)

    for chapter in range(1, 4):
        doc.add_heading(f"Chapter {chapter}", 1)

        for section in range(1, 4):
            doc.add_heading(f"Section {chapter}.{section}", 2)

            # Add several paragraphs
            for para in range(1, 4):
                doc.add_paragraph(
                    f"This is paragraph {para} of section {chapter}.{section}. "
                    f"It contains some text to make the document more realistic. "
                    f"Lorem ipsum dolor sit amet, consectetur adipiscing elit."
                )

            # Add a small table in some sections
            if section == 2:
                table = doc.add_table(rows=2, cols=2)
                table.rows[0].cells[0].text = f"Chapter {chapter}"
                table.rows[0].cells[1].text = f"Section {section}"
                table.rows[1].cells[0].text = "Data"
                table.rows[1].cells[1].text = "Value"

    return doc


def create_numbering_document() -> Document:
    """Create a document with programmatically-defined bullet and numbered lists.

    This exercises the numbering creation API: creating abstract numbering definitions,
    adding levels, creating num instances, and applying numbering to paragraphs.
    """
    doc = Document()

    doc.add_heading("Test Document - Custom Numbering", 0)

    numbering_part = doc.part.numbering_part

    # --- Bullet list ---
    an_bullet = numbering_part.add_abstract_num("singleLevel")
    an_bullet.add_lvl(
        ilvl=0,
        num_fmt="bullet",
        lvl_text="\uF0B7",
        indent_left=720,
        indent_hanging=360,
        font_name="Symbol",
    )
    num_bullet = numbering_part.add_num(an_bullet.abstractNumId)

    doc.add_heading("Bullet List", 1)
    for text in ["First bullet item", "Second bullet item", "Third bullet item"]:
        p = doc.add_paragraph(text)
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num_bullet.numId
        numPr.get_or_add_ilvl().val = 0

    # --- Decimal numbered list ---
    an_decimal = numbering_part.add_abstract_num("singleLevel")
    an_decimal.add_lvl(
        ilvl=0,
        num_fmt="decimal",
        lvl_text="%1.",
        indent_left=720,
        indent_hanging=360,
    )
    num_decimal = numbering_part.add_num(an_decimal.abstractNumId)

    doc.add_heading("Numbered List", 1)
    for text in ["First numbered item", "Second numbered item", "Third numbered item"]:
        p = doc.add_paragraph(text)
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num_decimal.numId
        numPr.get_or_add_ilvl().val = 0

    # --- Multi-level list ---
    an_multi = numbering_part.add_abstract_num("multiLevel")
    an_multi.add_lvl(
        ilvl=0, num_fmt="decimal", lvl_text="%1.",
        indent_left=720, indent_hanging=360,
    )
    an_multi.add_lvl(
        ilvl=1, num_fmt="lowerLetter", lvl_text="%2.",
        indent_left=1440, indent_hanging=360,
    )
    an_multi.add_lvl(
        ilvl=2, num_fmt="lowerRoman", lvl_text="%3.",
        indent_left=2160, indent_hanging=360,
    )
    num_multi = numbering_part.add_num(an_multi.abstractNumId)

    doc.add_heading("Multi-Level List", 1)
    items = [
        (0, "Top level item"),
        (1, "Sub-item a"),
        (1, "Sub-item b"),
        (2, "Sub-sub-item i"),
        (0, "Another top level"),
    ]
    for ilvl, text in items:
        p = doc.add_paragraph(text)
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num_multi.numId
        numPr.get_or_add_ilvl().val = ilvl

    return doc


# Dictionary of all generators for easy access
DOCUMENT_GENERATORS = {
    "basic": create_basic_document,
    "table": create_table_document,
    "nested_table": create_nested_table_document,
    "styled": create_styled_document,
    "multi_section": create_multi_section_document,
    "header_footer": create_header_footer_document,
    "font_formatting": create_font_formatting_document,
    "mixed_content": create_mixed_content_document,
    "unicode": create_unicode_document,
    "long_document": create_long_document,
    "numbering": create_numbering_document,
}
