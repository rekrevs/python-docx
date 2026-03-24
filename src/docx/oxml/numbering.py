"""Custom element classes related to the numbering part."""

from __future__ import annotations

import random
from typing import TYPE_CHECKING, Callable, List, cast

from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.shared import CT_DecimalNumber, CT_String
from docx.oxml.simpletypes import ST_DecimalNumber
from docx.oxml.xmlchemy import (
    BaseOxmlElement,
    OneAndOnlyOne,
    RequiredAttribute,
    ZeroOrMore,
    ZeroOrOne,
)

if TYPE_CHECKING:
    from docx.oxml.text.parfmt import CT_PPr
    from docx.oxml.text.run import CT_RPr


class CT_Lvl(BaseOxmlElement):
    """`w:lvl` element, which specifies the formatting and content of a particular level
    in a numbering definition.

    Schema children in order:
      start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText,
      lvlPicBulletId, legacy, lvlJc, pPr, rPr
    """

    # -- children (in schema order) --
    start: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:start",
        successors=(
            "w:numFmt",
            "w:lvlRestart",
            "w:pStyle",
            "w:isLgl",
            "w:suff",
            "w:lvlText",
            "w:lvlPicBulletId",
            "w:legacy",
            "w:lvlJc",
            "w:pPr",
            "w:rPr",
        ),
    )
    numFmt: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numFmt",
        successors=(
            "w:lvlRestart",
            "w:pStyle",
            "w:isLgl",
            "w:suff",
            "w:lvlText",
            "w:lvlPicBulletId",
            "w:legacy",
            "w:lvlJc",
            "w:pPr",
            "w:rPr",
        ),
    )
    lvlText: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lvlText",
        successors=(
            "w:lvlPicBulletId",
            "w:legacy",
            "w:lvlJc",
            "w:pPr",
            "w:rPr",
        ),
    )
    lvlJc: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:lvlJc",
        successors=(
            "w:pPr",
            "w:rPr",
        ),
    )
    pPr: CT_PPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:pPr",
        successors=("w:rPr",),
    )
    rPr: CT_RPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:rPr",
        successors=(),
    )

    # -- attributes --
    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    @classmethod
    def new(
        cls,
        ilvl: int,
        num_fmt: str,
        lvl_text: str,
        lvl_jc: str = "left",
        start_val: int = 1,
        indent_left: int | None = None,
        indent_hanging: int | None = None,
        font_name: str | None = None,
    ) -> CT_Lvl:
        """Create a new `w:lvl` element with the specified configuration.

        Args:
            ilvl: The level index (0-8).
            num_fmt: Number format (e.g. "decimal", "bullet", "lowerLetter").
            lvl_text: Level text pattern (e.g. "%1.", bullet char).
            lvl_jc: Justification ("left", "center", "right"). Defaults to "left".
            start_val: Starting value for this level. Defaults to 1.
            indent_left: Left indent in twips, or None for no indent.
            indent_hanging: Hanging indent in twips, or None.
            font_name: Font name for bullet characters (e.g. "Symbol"), or None.

        Returns:
            A new CT_Lvl element with the specified children.
        """
        lvl = cast(CT_Lvl, OxmlElement("w:lvl"))
        lvl.set(qn("w:ilvl"), str(ilvl))

        # w:start
        start_elm = CT_DecimalNumber.new("w:start", start_val)
        lvl.append(start_elm)

        # w:numFmt
        num_fmt_elm = CT_String.new("w:numFmt", num_fmt)
        lvl.append(num_fmt_elm)

        # w:lvlText
        lvl_text_elm = cast(CT_String, OxmlElement("w:lvlText"))
        lvl_text_elm.set(qn("w:val"), lvl_text)
        lvl.append(lvl_text_elm)

        # w:lvlJc
        lvl_jc_elm = CT_String.new("w:lvlJc", lvl_jc)
        lvl.append(lvl_jc_elm)

        # w:pPr with w:ind (if indent specified)
        if indent_left is not None:
            pPr_elm = OxmlElement("w:pPr")
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), str(indent_left))
            if indent_hanging is not None:
                ind.set(qn("w:hanging"), str(indent_hanging))
            pPr_elm.append(ind)
            lvl.append(pPr_elm)

        # w:rPr with w:rFonts (for bullet characters)
        if font_name is not None:
            rPr_elm = OxmlElement("w:rPr")
            rFonts = OxmlElement("w:rFonts")
            rFonts.set(qn("w:ascii"), font_name)
            rFonts.set(qn("w:hAnsi"), font_name)
            rFonts.set(qn("w:hint"), "default")
            rPr_elm.append(rFonts)
            lvl.append(rPr_elm)

        return lvl

    @property
    def start_val(self) -> int | None:
        """The value of the `w:start` child's `val` attribute, or None if no start child."""
        start_elm = self.start
        if start_elm is None:
            return None
        return start_elm.val

    @property
    def numFmt_val(self) -> str | None:
        """The value of the `w:numFmt` child's `val` attribute, or None if absent."""
        numFmt_elm = self.numFmt
        if numFmt_elm is None:
            return None
        return numFmt_elm.val

    @property
    def lvlText_val(self) -> str | None:
        """The value of the `w:lvlText` child's `val` attribute, or None if absent."""
        lvl_text_elm = self.lvlText
        if lvl_text_elm is None:
            return None
        return lvl_text_elm.get(qn("w:val"))

    @property
    def lvlJc_val(self) -> str | None:
        """The value of the `w:lvlJc` child's `val` attribute, or None if absent."""
        lvl_jc_elm = self.lvlJc
        if lvl_jc_elm is None:
            return None
        return lvl_jc_elm.val


class CT_AbstractNum(BaseOxmlElement):
    """`w:abstractNum` element, which defines an abstract numbering definition containing
    the formatting and content of each level in a multi-level numbering scheme.

    Schema children in order:
      nsid, multiLevelType, tmpl, name, styleLink, numStyleLink, lvl* (0..9)
    """

    # -- children (in schema order) --
    nsid: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:nsid",
        successors=(
            "w:multiLevelType",
            "w:tmpl",
            "w:name",
            "w:styleLink",
            "w:numStyleLink",
            "w:lvl",
        ),
    )
    multiLevelType: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:multiLevelType",
        successors=(
            "w:tmpl",
            "w:name",
            "w:styleLink",
            "w:numStyleLink",
            "w:lvl",
        ),
    )
    tmpl: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:tmpl",
        successors=(
            "w:name",
            "w:styleLink",
            "w:numStyleLink",
            "w:lvl",
        ),
    )
    lvl = ZeroOrMore(
        "w:lvl",
        successors=(),
    )

    # -- type-declarations for methods added by metaclass --
    lvl_lst: List[CT_Lvl]
    _add_lvl: Callable[..., CT_Lvl]
    _insert_lvl: Callable[[CT_Lvl], CT_Lvl]

    # -- attributes --
    abstractNumId: int = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId", ST_DecimalNumber
    )

    @classmethod
    def new(
        cls,
        abstract_num_id: int,
        multi_level_type: str = "hybridMultilevel",
    ) -> CT_AbstractNum:
        """Create a new `w:abstractNum` element with required children.

        Args:
            abstract_num_id: The unique ID for this abstract numbering definition.
            multi_level_type: One of "singleLevel", "multiLevel", "hybridMultilevel".
                Defaults to "hybridMultilevel".

        Returns:
            A new CT_AbstractNum element with nsid, multiLevelType, and tmpl children.
        """
        abstract_num = cast(CT_AbstractNum, OxmlElement("w:abstractNum"))
        abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

        # w:nsid -- random unique hex identifier
        nsid_elm = CT_String.new("w:nsid", _generate_nsid())
        abstract_num.append(nsid_elm)

        # w:multiLevelType
        mlt = CT_String.new("w:multiLevelType", multi_level_type)
        abstract_num.append(mlt)

        # w:tmpl -- random template ID
        tmpl_elm = CT_String.new("w:tmpl", _generate_nsid())
        abstract_num.append(tmpl_elm)

        return abstract_num

    def add_lvl(
        self,
        ilvl: int,
        num_fmt: str,
        lvl_text: str,
        lvl_jc: str = "left",
        start_val: int = 1,
        indent_left: int | None = None,
        indent_hanging: int | None = None,
        font_name: str | None = None,
    ) -> CT_Lvl:
        """Add a new `w:lvl` child element with the specified configuration.

        Args:
            ilvl: The level index (0-8).
            num_fmt: Number format string.
            lvl_text: Level text pattern.
            lvl_jc: Justification. Defaults to "left".
            start_val: Starting value. Defaults to 1.
            indent_left: Left indent in twips, or None.
            indent_hanging: Hanging indent in twips, or None.
            font_name: Font name for bullets, or None.

        Returns:
            The newly created CT_Lvl element.
        """
        new_lvl = CT_Lvl.new(
            ilvl=ilvl,
            num_fmt=num_fmt,
            lvl_text=lvl_text,
            lvl_jc=lvl_jc,
            start_val=start_val,
            indent_left=indent_left,
            indent_hanging=indent_hanging,
            font_name=font_name,
        )
        return self._insert_lvl(new_lvl)

    @property
    def multiLevelType_val(self) -> str | None:
        """The value of the `w:multiLevelType` child's `val` attribute, or None."""
        mlt = self.multiLevelType
        if mlt is None:
            return None
        return mlt.val


class CT_Num(BaseOxmlElement):
    """``<w:num>`` element, which represents a concrete list definition instance, having
    a required child <w:abstractNumId> that references an abstract numbering definition
    that defines most of the formatting details."""

    abstractNumId: CT_DecimalNumber = OneAndOnlyOne(  # pyright: ignore[reportAssignmentType]
        "w:abstractNumId"
    )
    lvlOverride = ZeroOrMore("w:lvlOverride")
    numId: int = RequiredAttribute("w:numId", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    # -- type-declarations for methods added by metaclass --
    _add_lvlOverride: Callable[..., CT_NumLvl]  # noqa: F821
    _insert_num: Callable[[CT_Num], CT_Num]

    def add_lvlOverride(self, ilvl: int) -> CT_NumLvl:  # noqa: F821
        """Return a newly added CT_NumLvl (<w:lvlOverride>) element having its ``ilvl``
        attribute set to `ilvl`."""
        return self._add_lvlOverride(ilvl=ilvl)

    @classmethod
    def new(cls, num_id: int, abstractNum_id: int) -> CT_Num:
        """Return a new ``<w:num>`` element having numId of `num_id` and having a
        ``<w:abstractNumId>`` child with val attribute set to `abstractNum_id`."""
        num = cast(CT_Num, OxmlElement("w:num"))
        num.set(qn("w:numId"), str(num_id))
        abstractNumId_elm = CT_DecimalNumber.new("w:abstractNumId", abstractNum_id)
        num.append(abstractNumId_elm)
        return num


class CT_NumLvl(BaseOxmlElement):
    """``<w:lvlOverride>`` element, which identifies a level in a list definition to
    override with settings it contains."""

    startOverride: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:startOverride", successors=("w:lvl",)
    )
    ilvl: int = RequiredAttribute("w:ilvl", ST_DecimalNumber)  # pyright: ignore[reportAssignmentType]

    # -- type-declarations for methods added by metaclass --
    _add_startOverride: Callable[..., CT_DecimalNumber]

    def add_startOverride(self, val: int) -> CT_DecimalNumber:
        """Return a newly added CT_DecimalNumber element having tagname
        ``w:startOverride`` and ``val`` attribute set to `val`."""
        return self._add_startOverride(val=val)


class CT_NumPr(BaseOxmlElement):
    """A ``<w:numPr>`` element, a container for numbering properties applied to a
    paragraph."""

    ilvl: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:ilvl", successors=("w:numId", "w:numberingChange", "w:ins")
    )
    numId: CT_DecimalNumber | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:numId", successors=("w:numberingChange", "w:ins")
    )


class CT_Numbering(BaseOxmlElement):
    """``<w:numbering>`` element, the root element of a numbering part, i.e.
    numbering.xml."""

    abstractNum = ZeroOrMore(
        "w:abstractNum", successors=("w:num", "w:numIdMacAtCleanup")
    )
    num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))

    # -- type-declarations for methods added by metaclass --
    abstractNum_lst: List[CT_AbstractNum]
    num_lst: List[CT_Num]
    _add_abstractNum: Callable[..., CT_AbstractNum]
    _insert_abstractNum: Callable[[CT_AbstractNum], CT_AbstractNum]
    _insert_num: Callable[[CT_Num], CT_Num]

    def add_abstractNum(
        self,
        multi_level_type: str = "hybridMultilevel",
    ) -> CT_AbstractNum:
        """Return a newly added CT_AbstractNum element with auto-assigned abstractNumId.

        Args:
            multi_level_type: One of "singleLevel", "multiLevel", "hybridMultilevel".

        Returns:
            The newly created and inserted CT_AbstractNum element.
        """
        next_id = self._next_abstractNumId
        abstract_num = CT_AbstractNum.new(next_id, multi_level_type)
        return self._insert_abstractNum(abstract_num)

    def add_num(self, abstractNum_id: int) -> CT_Num:
        """Return a newly added CT_Num (<w:num>) element referencing the abstract
        numbering definition identified by `abstractNum_id`."""
        next_num_id = self._next_numId
        num = CT_Num.new(next_num_id, abstractNum_id)
        return self._insert_num(num)

    def num_having_numId(self, numId: int) -> CT_Num:
        """Return the ``<w:num>`` child element having ``numId`` attribute matching
        `numId`."""
        xpath = './w:num[@w:numId="%d"]' % numId
        try:
            return self.xpath(xpath)[0]
        except IndexError:
            raise KeyError("no <w:num> element with numId %d" % numId)

    @property
    def _next_abstractNumId(self) -> int:
        """The first ``abstractNumId`` unused by a ``<w:abstractNum>`` element,
        starting at 0 and filling any gaps."""
        id_strs = self.xpath("./w:abstractNum/@w:abstractNumId")
        ids = [int(id_str) for id_str in id_strs]
        for candidate in range(len(ids) + 1):
            if candidate not in ids:
                return candidate
        return len(ids)  # pragma: no cover

    @property
    def _next_numId(self) -> int:
        """The first ``numId`` unused by a ``<w:num>`` element, starting at 1 and
        filling any gaps in numbering between existing ``<w:num>`` elements."""
        numId_strs = self.xpath("./w:num/@w:numId")
        num_ids = [int(numId_str) for numId_str in numId_strs]
        for num in range(1, len(num_ids) + 2):
            if num not in num_ids:
                return num
        return len(num_ids) + 1  # pragma: no cover


def _generate_nsid() -> str:
    """Generate a random 8-character uppercase hex string for use as w:nsid or w:tmpl value.

    The format matches the OOXML standard (e.g., "FFFFFF7C", "3D1EFFD4").
    """
    return "%08X" % random.randint(0, 0xFFFFFFFF)
