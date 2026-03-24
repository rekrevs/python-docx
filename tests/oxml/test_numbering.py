# pyright: reportPrivateUsage=false

"""Unit-test suite for `docx.oxml.numbering` module."""

from __future__ import annotations

from typing import cast

import pytest

from docx.oxml.ns import qn
from docx.oxml.numbering import CT_AbstractNum, CT_Lvl, CT_Numbering

from ..unitutil.cxml import element


class DescribeCT_Lvl:
    """Unit-test suite for `docx.oxml.numbering.CT_Lvl`."""

    def it_can_create_a_new_decimal_level(self):
        lvl = CT_Lvl.new(ilvl=0, num_fmt="decimal", lvl_text="%1.", lvl_jc="left")

        assert lvl.ilvl == 0
        assert lvl.start_val == 1
        assert lvl.numFmt_val == "decimal"
        assert lvl.lvlText_val == "%1."
        assert lvl.lvlJc_val == "left"
        # No pPr or rPr by default
        assert lvl.pPr is None
        assert lvl.rPr is None

    def it_can_create_a_bullet_level_with_font(self):
        lvl = CT_Lvl.new(
            ilvl=0,
            num_fmt="bullet",
            lvl_text="\uF0B7",
            lvl_jc="left",
            start_val=1,
            indent_left=720,
            indent_hanging=360,
            font_name="Symbol",
        )

        assert lvl.ilvl == 0
        assert lvl.numFmt_val == "bullet"
        assert lvl.lvlText_val == "\uF0B7"
        # Check pPr/ind
        pPr = lvl.pPr
        assert pPr is not None
        ind = pPr.find(qn("w:ind"))
        assert ind is not None
        assert ind.get(qn("w:left")) == "720"
        assert ind.get(qn("w:hanging")) == "360"
        # Check rPr/rFonts
        rPr = lvl.rPr
        assert rPr is not None
        rFonts = rPr.find(qn("w:rFonts"))
        assert rFonts is not None
        assert rFonts.get(qn("w:ascii")) == "Symbol"
        assert rFonts.get(qn("w:hAnsi")) == "Symbol"

    def it_can_create_a_level_with_custom_start_value(self):
        lvl = CT_Lvl.new(ilvl=2, num_fmt="lowerRoman", lvl_text="%3)", start_val=5)

        assert lvl.ilvl == 2
        assert lvl.start_val == 5
        assert lvl.numFmt_val == "lowerRoman"
        assert lvl.lvlText_val == "%3)"

    def it_can_create_a_level_with_indent_but_no_font(self):
        lvl = CT_Lvl.new(
            ilvl=1,
            num_fmt="decimal",
            lvl_text="%2.",
            indent_left=1440,
            indent_hanging=360,
        )

        assert lvl.pPr is not None
        assert lvl.rPr is None

    def it_returns_None_for_missing_children_properties(self):
        # A minimal lvl element without the optional children
        lvl_elm = cast(CT_Lvl, element("w:lvl{w:ilvl=0}"))

        assert lvl_elm.start_val is None
        assert lvl_elm.numFmt_val is None
        assert lvl_elm.lvlText_val is None
        assert lvl_elm.lvlJc_val is None


class DescribeCT_AbstractNum:
    """Unit-test suite for `docx.oxml.numbering.CT_AbstractNum`."""

    def it_can_create_a_new_abstract_num(self):
        abstract_num = CT_AbstractNum.new(abstract_num_id=0)

        assert abstract_num.abstractNumId == 0
        assert abstract_num.multiLevelType_val == "hybridMultilevel"
        # nsid should be present (8-char hex)
        nsid = abstract_num.nsid
        assert nsid is not None
        assert len(nsid.val) == 8
        # tmpl should be present
        tmpl = abstract_num.tmpl
        assert tmpl is not None
        assert len(tmpl.val) == 8

    def it_can_create_a_single_level_abstract_num(self):
        abstract_num = CT_AbstractNum.new(
            abstract_num_id=5,
            multi_level_type="singleLevel",
        )

        assert abstract_num.abstractNumId == 5
        assert abstract_num.multiLevelType_val == "singleLevel"

    def it_can_add_levels(self):
        abstract_num = CT_AbstractNum.new(abstract_num_id=0)

        lvl0 = abstract_num.add_lvl(
            ilvl=0,
            num_fmt="decimal",
            lvl_text="%1.",
            indent_left=720,
            indent_hanging=360,
        )
        lvl1 = abstract_num.add_lvl(
            ilvl=1,
            num_fmt="lowerLetter",
            lvl_text="%2.",
            indent_left=1440,
            indent_hanging=360,
        )

        assert len(abstract_num.lvl_lst) == 2
        assert abstract_num.lvl_lst[0].ilvl == 0
        assert abstract_num.lvl_lst[1].ilvl == 1
        assert lvl0.numFmt_val == "decimal"
        assert lvl1.numFmt_val == "lowerLetter"

    def it_can_add_a_bullet_level_with_font(self):
        abstract_num = CT_AbstractNum.new(abstract_num_id=0, multi_level_type="singleLevel")

        lvl = abstract_num.add_lvl(
            ilvl=0,
            num_fmt="bullet",
            lvl_text="\uF0B7",
            indent_left=720,
            indent_hanging=360,
            font_name="Symbol",
        )

        assert lvl.numFmt_val == "bullet"
        rPr = lvl.rPr
        assert rPr is not None
        rFonts = rPr.find(qn("w:rFonts"))
        assert rFonts is not None
        assert rFonts.get(qn("w:ascii")) == "Symbol"

    def it_returns_None_for_missing_multiLevelType(self):
        # A bare abstractNum element without children
        abstract_num = cast(CT_AbstractNum, element("w:abstractNum{w:abstractNumId=0}"))
        assert abstract_num.multiLevelType_val is None

    def it_can_add_nine_levels(self):
        """OOXML supports up to 9 levels (ilvl 0-8)."""
        abstract_num = CT_AbstractNum.new(abstract_num_id=0, multi_level_type="multiLevel")

        for i in range(9):
            abstract_num.add_lvl(
                ilvl=i,
                num_fmt="decimal",
                lvl_text=f"%{i+1}.",
                indent_left=720 * (i + 1),
                indent_hanging=360,
            )

        assert len(abstract_num.lvl_lst) == 9
        for i, lvl in enumerate(abstract_num.lvl_lst):
            assert lvl.ilvl == i
            assert lvl.lvlText_val == f"%{i+1}."


class DescribeCT_Numbering:
    """Unit-test suite for `docx.oxml.numbering.CT_Numbering` extensions."""

    def it_can_add_an_abstract_num(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        abstract_num = numbering.add_abstractNum("singleLevel")

        assert abstract_num.abstractNumId == 0
        assert abstract_num.multiLevelType_val == "singleLevel"
        assert len(numbering.abstractNum_lst) == 1

    def it_auto_assigns_abstract_num_ids(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        an0 = numbering.add_abstractNum()
        an1 = numbering.add_abstractNum()
        an2 = numbering.add_abstractNum()

        assert an0.abstractNumId == 0
        assert an1.abstractNumId == 1
        assert an2.abstractNumId == 2

    @pytest.mark.parametrize(
        ("cxml", "expected_next_id"),
        [
            ("w:numbering", 0),
            ("w:numbering/w:abstractNum{w:abstractNumId=0}", 1),
            (
                "w:numbering/(w:abstractNum{w:abstractNumId=0}"
                ",w:abstractNum{w:abstractNumId=1})",
                2,
            ),
            (
                "w:numbering/(w:abstractNum{w:abstractNumId=0}"
                ",w:abstractNum{w:abstractNumId=2})",
                1,
            ),
        ],
    )
    def it_computes_the_next_abstractNumId(self, cxml: str, expected_next_id: int):
        numbering = cast(CT_Numbering, element(cxml))
        assert numbering._next_abstractNumId == expected_next_id

    def it_inserts_abstractNum_before_num_elements(self):
        """Schema ordering: w:abstractNum must precede w:num."""
        numbering = cast(CT_Numbering, element("w:numbering"))

        # Add a num first
        num = numbering.add_num(0)
        # Then add an abstractNum -- it should be inserted before the num
        abstract_num = numbering.add_abstractNum()

        children = list(numbering)
        abstract_num_idx = children.index(abstract_num)
        num_idx = children.index(num)
        assert abstract_num_idx < num_idx

    def it_can_add_num_referencing_abstract_num(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        abstract_num = numbering.add_abstractNum()
        num = numbering.add_num(abstract_num.abstractNumId)

        assert num.numId == 1
        assert num.abstractNumId.val == 0

    def it_can_create_multiple_definitions(self):
        numbering = cast(CT_Numbering, element("w:numbering"))

        an0 = numbering.add_abstractNum("singleLevel")
        an0.add_lvl(ilvl=0, num_fmt="bullet", lvl_text="\uF0B7", font_name="Symbol")
        num0 = numbering.add_num(an0.abstractNumId)

        an1 = numbering.add_abstractNum("singleLevel")
        an1.add_lvl(ilvl=0, num_fmt="decimal", lvl_text="%1.")
        num1 = numbering.add_num(an1.abstractNumId)

        assert len(numbering.abstractNum_lst) == 2
        assert len(numbering.num_lst) == 2
        assert num0.numId == 1
        assert num1.numId == 2
