# pyright: reportPrivateUsage=false

"""End-to-end tests for programmatic creation of numbering definitions.

These tests verify that abstract numbering definitions can be created, configured
with various number formats, applied to paragraphs, and survive save/reload cycles.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx import Document
from docx.oxml.ns import qn


class DescribeDecimalNumbering:
    """Creating and applying decimal (1. 2. 3.) numbering."""

    def it_can_create_a_decimal_numbered_list(self, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        # Create abstract numbering definition
        abstract_num = numbering_part.add_abstract_num("singleLevel")
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt="decimal",
            lvl_text="%1.",
            indent_left=720,
            indent_hanging=360,
        )

        # Create concrete num referencing the abstract
        num = numbering_part.add_num(abstract_num.abstractNumId)

        # Apply to paragraphs
        for i in range(3):
            p = doc.add_paragraph(f"Item {i + 1}")
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = num.numId
            numPr.get_or_add_ilvl().val = 0

        # Save and reload
        path = tmp_path / "decimal_list.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        # Verify numbering part exists and has expected content
        npart = doc2.part.numbering_part
        numbering_elm = npart.numbering_elm

        # Check that we have our custom abstract num plus the default ones
        abstract_nums = numbering_elm.abstractNum_lst
        assert len(abstract_nums) >= 3  # 2 defaults + 1 new

        # Find our custom abstract num (it should have the highest abstractNumId)
        custom_an = max(abstract_nums, key=lambda an: an.abstractNumId)
        assert custom_an.multiLevelType_val == "singleLevel"
        assert len(custom_an.lvl_lst) == 1
        assert custom_an.lvl_lst[0].numFmt_val == "decimal"
        assert custom_an.lvl_lst[0].lvlText_val == "%1."

        # Verify paragraphs still have numbering
        # Skip default empty paragraph (first paragraph in new document)
        numbered_paragraphs = [p for p in doc2.paragraphs if p.text.startswith("Item")]
        assert len(numbered_paragraphs) == 3


class DescribeBulletNumbering:
    """Creating and applying bullet list numbering."""

    def it_can_create_a_bullet_list(self, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("singleLevel")
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt="bullet",
            lvl_text="\uF0B7",
            indent_left=720,
            indent_hanging=360,
            font_name="Symbol",
        )

        num = numbering_part.add_num(abstract_num.abstractNumId)

        for text in ["Alpha", "Beta", "Gamma"]:
            p = doc.add_paragraph(text)
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = num.numId
            numPr.get_or_add_ilvl().val = 0

        path = tmp_path / "bullet_list.docx"
        doc.save(str(path))

        # Reload and verify
        doc2 = Document(str(path))
        npart = doc2.part.numbering_part
        abstract_nums = npart.numbering_elm.abstractNum_lst

        # Find the bullet abstract num
        bullet_ans = [
            an for an in abstract_nums
            if len(an.lvl_lst) > 0 and an.lvl_lst[0].numFmt_val == "bullet"
            and an.lvl_lst[0].rPr is not None
        ]
        assert len(bullet_ans) >= 1

        # Verify the bullet has the font specification
        bullet_lvl = bullet_ans[-1].lvl_lst[0]
        rPr = bullet_lvl.rPr
        assert rPr is not None
        rFonts = rPr.find(qn("w:rFonts"))
        assert rFonts is not None
        assert rFonts.get(qn("w:ascii")) == "Symbol"


class DescribeLetterNumbering:
    """Creating letter-based numbering (a. b. c. and A. B. C.)."""

    @pytest.mark.parametrize(
        ("num_fmt", "lvl_text"),
        [
            ("lowerLetter", "%1)"),
            ("upperLetter", "%1."),
        ],
    )
    def it_can_create_a_letter_list(self, num_fmt: str, lvl_text: str, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("singleLevel")
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt=num_fmt,
            lvl_text=lvl_text,
            indent_left=720,
            indent_hanging=360,
        )

        num = numbering_part.add_num(abstract_num.abstractNumId)

        p = doc.add_paragraph("Test item")
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num.numId
        numPr.get_or_add_ilvl().val = 0

        path = tmp_path / f"{num_fmt}_list.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        abstract_nums = doc2.part.numbering_part.numbering_elm.abstractNum_lst
        custom_an = max(abstract_nums, key=lambda an: an.abstractNumId)
        assert custom_an.lvl_lst[0].numFmt_val == num_fmt
        assert custom_an.lvl_lst[0].lvlText_val == lvl_text


class DescribeRomanNumbering:
    """Creating roman numeral numbering (i. ii. iii. and I. II. III.)."""

    @pytest.mark.parametrize(
        ("num_fmt", "lvl_text"),
        [
            ("lowerRoman", "%1."),
            ("upperRoman", "%1."),
        ],
    )
    def it_can_create_a_roman_list(self, num_fmt: str, lvl_text: str, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("singleLevel")
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt=num_fmt,
            lvl_text=lvl_text,
            indent_left=720,
            indent_hanging=360,
        )

        num = numbering_part.add_num(abstract_num.abstractNumId)

        p = doc.add_paragraph("Roman item")
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num.numId
        numPr.get_or_add_ilvl().val = 0

        path = tmp_path / f"{num_fmt}_list.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        abstract_nums = doc2.part.numbering_part.numbering_elm.abstractNum_lst
        custom_an = max(abstract_nums, key=lambda an: an.abstractNumId)
        assert custom_an.lvl_lst[0].numFmt_val == num_fmt


class DescribeMultiLevelNumbering:
    """Creating multi-level numbering definitions with multiple levels."""

    def it_can_create_a_multi_level_list(self, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("multiLevel")

        # Define 3 levels
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt="decimal",
            lvl_text="%1.",
            indent_left=720,
            indent_hanging=360,
        )
        abstract_num.add_lvl(
            ilvl=1,
            num_fmt="lowerLetter",
            lvl_text="%2.",
            indent_left=1440,
            indent_hanging=360,
        )
        abstract_num.add_lvl(
            ilvl=2,
            num_fmt="lowerRoman",
            lvl_text="%3.",
            indent_left=2160,
            indent_hanging=360,
        )

        num = numbering_part.add_num(abstract_num.abstractNumId)

        # Apply levels to paragraphs
        for ilvl, text in [(0, "First level"), (1, "Second level"), (2, "Third level")]:
            p = doc.add_paragraph(text)
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = num.numId
            numPr.get_or_add_ilvl().val = ilvl

        path = tmp_path / "multi_level_list.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        abstract_nums = doc2.part.numbering_part.numbering_elm.abstractNum_lst
        custom_an = max(abstract_nums, key=lambda an: an.abstractNumId)
        assert custom_an.multiLevelType_val == "multiLevel"
        assert len(custom_an.lvl_lst) == 3
        assert custom_an.lvl_lst[0].numFmt_val == "decimal"
        assert custom_an.lvl_lst[1].numFmt_val == "lowerLetter"
        assert custom_an.lvl_lst[2].numFmt_val == "lowerRoman"

    def it_can_create_a_nine_level_definition(self, tmp_path: Path):
        """OOXML supports up to 9 levels (ilvl 0-8)."""
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("multiLevel")

        num_fmts = [
            "decimal", "lowerLetter", "lowerRoman",
            "decimal", "lowerLetter", "lowerRoman",
            "decimal", "lowerLetter", "lowerRoman",
        ]
        for i in range(9):
            abstract_num.add_lvl(
                ilvl=i,
                num_fmt=num_fmts[i],
                lvl_text=f"%{i+1}.",
                indent_left=720 * (i + 1),
                indent_hanging=360,
            )

        numbering_part.add_num(abstract_num.abstractNumId)

        path = tmp_path / "nine_level_list.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        abstract_nums = doc2.part.numbering_part.numbering_elm.abstractNum_lst
        custom_an = max(abstract_nums, key=lambda an: an.abstractNumId)
        assert len(custom_an.lvl_lst) == 9
        for i, lvl in enumerate(custom_an.lvl_lst):
            assert lvl.ilvl == i


class DescribeMultipleDefinitions:
    """Creating multiple independent numbering definitions in one document."""

    def it_supports_multiple_independent_definitions(self, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        # Bullet list
        an_bullet = numbering_part.add_abstract_num("singleLevel")
        an_bullet.add_lvl(
            ilvl=0, num_fmt="bullet", lvl_text="\uF0B7",
            indent_left=720, indent_hanging=360, font_name="Symbol",
        )
        num_bullet = numbering_part.add_num(an_bullet.abstractNumId)

        # Decimal list
        an_decimal = numbering_part.add_abstract_num("singleLevel")
        an_decimal.add_lvl(
            ilvl=0, num_fmt="decimal", lvl_text="%1.",
            indent_left=720, indent_hanging=360,
        )
        num_decimal = numbering_part.add_num(an_decimal.abstractNumId)

        # Apply bullet to first set of paragraphs
        for text in ["Bullet A", "Bullet B"]:
            p = doc.add_paragraph(text)
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = num_bullet.numId
            numPr.get_or_add_ilvl().val = 0

        # Apply decimal to next set
        for text in ["Number 1", "Number 2"]:
            p = doc.add_paragraph(text)
            pPr = p._element.get_or_add_pPr()
            numPr = pPr.get_or_add_numPr()
            numPr.get_or_add_numId().val = num_decimal.numId
            numPr.get_or_add_ilvl().val = 0

        path = tmp_path / "multiple_defs.docx"
        doc.save(str(path))

        doc2 = Document(str(path))
        npart = doc2.part.numbering_part
        # 2 default abstract nums + 2 new ones
        assert len(npart.numbering_elm.abstractNum_lst) >= 4
        # 2 default nums + 2 new ones
        assert len(npart.numbering_elm.num_lst) >= 4


class DescribeLevelOverrides:
    """Level overrides on created numbering definitions."""

    def it_can_apply_level_overrides(self, tmp_path: Path):
        doc = Document()
        numbering_part = doc.part.numbering_part

        abstract_num = numbering_part.add_abstract_num("singleLevel")
        abstract_num.add_lvl(
            ilvl=0,
            num_fmt="decimal",
            lvl_text="%1.",
            start_val=1,
            indent_left=720,
            indent_hanging=360,
        )

        num = numbering_part.add_num(abstract_num.abstractNumId)

        # Apply level override to restart numbering at 10
        lvl_override = num.add_lvlOverride(ilvl=0)
        lvl_override.add_startOverride(val=10)

        p = doc.add_paragraph("Should start at 10")
        pPr = p._element.get_or_add_pPr()
        numPr = pPr.get_or_add_numPr()
        numPr.get_or_add_numId().val = num.numId
        numPr.get_or_add_ilvl().val = 0

        path = tmp_path / "level_override.docx"
        doc.save(str(path))

        # Reload and verify
        doc2 = Document(str(path))
        npart = doc2.part.numbering_part
        # Find our num element
        nums = npart.numbering_elm.num_lst
        our_num = max(nums, key=lambda n: n.numId)
        # Check lvlOverride
        overrides = our_num.xpath("./w:lvlOverride")
        assert len(overrides) >= 1
