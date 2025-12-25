"""Custom element classes for Office Math Markup Language (OMML)."""

# pyright: reportUnknownMemberType=false, reportUnknownVariableType=false
# pyright: reportUnknownArgumentType=false, reportArgumentType=false
# pyright: reportIncompatibleMethodOverride=false, reportPrivateUsage=false
# pyright: reportAttributeAccessIssue=false, reportOperatorIssue=false

from __future__ import annotations

from typing import Any, List

from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import parse_xml
from docx.oxml.xmlchemy import BaseOxmlElement, ZeroOrMore, ZeroOrOne


class CT_OMath(BaseOxmlElement):
    """``<m:oMath>`` element - inline math zone.

    Contains mathematical content as child elements.
    """

    @classmethod
    def new(cls) -> CT_OMath:
        """Create a new empty m:oMath element."""
        return parse_xml(f"<m:oMath {nsdecls('m')}/>")  # type: ignore[return-value]

    @classmethod
    def new_with_text(cls, text: str) -> CT_OMath:
        """Create an m:oMath element containing simple text."""
        # Escape XML special characters
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:r><m:t>{text}</m:t></m:r>"
            f"</m:oMath>"
        )

    @classmethod
    def new_fraction(cls, numerator: str, denominator: str) -> CT_OMath:
        """Create an m:oMath element containing a fraction."""
        num = numerator.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        den = denominator.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:f>"
            f"<m:num><m:r><m:t>{num}</m:t></m:r></m:num>"
            f"<m:den><m:r><m:t>{den}</m:t></m:r></m:den>"
            f"</m:f>"
            f"</m:oMath>"
        )

    @classmethod
    def new_superscript(cls, base: str, superscript: str) -> CT_OMath:
        """Create an m:oMath element containing a superscript."""
        b = base.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        s = superscript.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:sSup>"
            f"<m:e><m:r><m:t>{b}</m:t></m:r></m:e>"
            f"<m:sup><m:r><m:t>{s}</m:t></m:r></m:sup>"
            f"</m:sSup>"
            f"</m:oMath>"
        )

    @classmethod
    def new_subscript(cls, base: str, subscript: str) -> CT_OMath:
        """Create an m:oMath element containing a subscript."""
        b = base.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        s = subscript.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:sSub>"
            f"<m:e><m:r><m:t>{b}</m:t></m:r></m:e>"
            f"<m:sub><m:r><m:t>{s}</m:t></m:r></m:sub>"
            f"</m:sSub>"
            f"</m:oMath>"
        )

    @classmethod
    def new_sqrt(cls, content: str) -> CT_OMath:
        """Create an m:oMath element containing a square root."""
        c = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:rad>"
            f'<m:radPr><m:degHide m:val="1"/></m:radPr>'
            f"<m:deg/>"
            f"<m:e><m:r><m:t>{c}</m:t></m:r></m:e>"
            f"</m:rad>"
            f"</m:oMath>"
        )

    @classmethod
    def new_nthroot(cls, content: str, degree: str) -> CT_OMath:
        """Create an m:oMath element containing an nth root."""
        c = content.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        d = degree.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return parse_xml(  # type: ignore[return-value]
            f'<m:oMath {nsdecls("m")}>'
            f"<m:rad>"
            f"<m:deg><m:r><m:t>{d}</m:t></m:r></m:deg>"
            f"<m:e><m:r><m:t>{c}</m:t></m:r></m:e>"
            f"</m:rad>"
            f"</m:oMath>"
        )

    @property
    def math_text(self) -> str:
        """Return plain text representation of the math content.

        This extracts all text from m:t elements within the math zone.
        """
        return _extract_math_text(self)


class CT_OMathPara(BaseOxmlElement):
    """``<m:oMathPara>`` element - math paragraph (block-level, often centered).

    Contains one or more m:oMath elements.
    """

    # Type annotation for metaclass-generated property
    oMath_lst: List[CT_OMath]

    oMath = ZeroOrMore("m:oMath")

    @property
    def math_text(self) -> str:
        """Return plain text representation of all math content."""
        texts: List[str] = []
        for omath in self.oMath_lst:
            texts.append(omath.math_text)
        return " ".join(texts)


class CT_MathRun(BaseOxmlElement):
    """``<m:r>`` element - math run containing text.

    Contains m:t (text) element with the actual characters.
    """

    _t: CT_MathText | None = ZeroOrOne("m:t")  # type: ignore[assignment]

    @property
    def run_text(self) -> str:
        """Return the text content of this math run."""
        t_elem = self._t
        if t_elem is not None and t_elem.text:
            return str(t_elem.text)
        return ""


class CT_MathText(BaseOxmlElement):
    """``<m:t>`` element - math text content."""

    pass  # Text content is in .text property inherited from lxml


# Helper functions for text extraction

def _extract_math_text(element: Any) -> str:
    """Recursively extract text from math elements.

    This handles the various math structures and extracts readable text.
    """
    texts = []

    for child in element:
        tag = child.tag
        local_name = tag.split("}")[-1] if "}" in tag else tag

        # Direct text content
        if local_name == "t":
            if child.text:
                texts.append(child.text)

        # Math run - contains text
        elif local_name == "r":
            t_elem = child.find(qn("m:t"))
            if t_elem is not None and t_elem.text:
                texts.append(t_elem.text)

        # Fraction: num/den
        elif local_name == "f":
            num = child.find(qn("m:num"))
            den = child.find(qn("m:den"))
            num_text = _extract_math_text(num) if num is not None else ""
            den_text = _extract_math_text(den) if den is not None else ""
            texts.append(f"({num_text}/{den_text})")

        # Superscript
        elif local_name == "sSup":
            e = child.find(qn("m:e"))
            sup = child.find(qn("m:sup"))
            e_text = _extract_math_text(e) if e is not None else ""
            sup_text = _extract_math_text(sup) if sup is not None else ""
            texts.append(f"{e_text}^{sup_text}")

        # Subscript
        elif local_name == "sSub":
            e = child.find(qn("m:e"))
            sub = child.find(qn("m:sub"))
            e_text = _extract_math_text(e) if e is not None else ""
            sub_text = _extract_math_text(sub) if sub is not None else ""
            texts.append(f"{e_text}_{sub_text}")

        # Sub-superscript
        elif local_name == "sSubSup":
            e = child.find(qn("m:e"))
            sub = child.find(qn("m:sub"))
            sup = child.find(qn("m:sup"))
            e_text = _extract_math_text(e) if e is not None else ""
            sub_text = _extract_math_text(sub) if sub is not None else ""
            sup_text = _extract_math_text(sup) if sup is not None else ""
            texts.append(f"{e_text}_{sub_text}^{sup_text}")

        # Radical (square root)
        elif local_name == "rad":
            deg = child.find(qn("m:deg"))
            e = child.find(qn("m:e"))
            e_text = _extract_math_text(e) if e is not None else ""
            if deg is not None:
                deg_text = _extract_math_text(deg)
                if deg_text:
                    texts.append(f"root[{deg_text}]({e_text})")
                else:
                    texts.append(f"sqrt({e_text})")
            else:
                texts.append(f"sqrt({e_text})")

        # N-ary (summation, product, integral)
        elif local_name == "nary":
            nary_pr = child.find(qn("m:naryPr"))
            chr_elem = nary_pr.find(qn("m:chr")) if nary_pr is not None else None
            chr_val = chr_elem.get(qn("m:val")) if chr_elem is not None else "∑"
            sub = child.find(qn("m:sub"))
            sup = child.find(qn("m:sup"))
            e = child.find(qn("m:e"))
            sub_text = _extract_math_text(sub) if sub is not None else ""
            sup_text = _extract_math_text(sup) if sup is not None else ""
            e_text = _extract_math_text(e) if e is not None else ""
            texts.append(f"{chr_val}[{sub_text},{sup_text}]({e_text})")

        # Delimiter (parentheses, brackets)
        elif local_name == "d":
            d_pr = child.find(qn("m:dPr"))
            beg_chr = "("
            end_chr = ")"
            if d_pr is not None:
                beg = d_pr.find(qn("m:begChr"))
                end = d_pr.find(qn("m:endChr"))
                if beg is not None:
                    beg_chr = beg.get(qn("m:val")) or "("
                if end is not None:
                    end_chr = end.get(qn("m:val")) or ")"
            e_elems = child.findall(qn("m:e"))
            inner_texts = [_extract_math_text(e) for e in e_elems]
            texts.append(f"{beg_chr}{', '.join(inner_texts)}{end_chr}")

        # Matrix
        elif local_name == "m":
            mr_elems = child.findall(qn("m:mr"))
            rows = []
            for mr in mr_elems:
                e_elems = mr.findall(qn("m:e"))
                row = [_extract_math_text(e) for e in e_elems]
                rows.append(", ".join(row))
            texts.append(f"[{'; '.join(rows)}]")

        # Equation array
        elif local_name == "eqArr":
            e_elems = child.findall(qn("m:e"))
            eq_texts = [_extract_math_text(e) for e in e_elems]
            texts.append(" ; ".join(eq_texts))

        # Function (sin, cos, etc.)
        elif local_name == "func":
            fname = child.find(qn("m:fName"))
            e = child.find(qn("m:e"))
            fname_text = _extract_math_text(fname) if fname is not None else "f"
            e_text = _extract_math_text(e) if e is not None else ""
            texts.append(f"{fname_text}({e_text})")

        # Other containers - recurse
        elif local_name in ("e", "num", "den", "sub", "sup", "deg", "fName",
                           "lim", "mr", "acc", "bar", "box", "borderBox",
                           "groupChr", "limLow", "limUpp", "phant", "sPre"):
            texts.append(_extract_math_text(child))

        # Skip property elements
        elif local_name.endswith("Pr"):
            continue

        # Unknown - try to recurse
        else:
            inner = _extract_math_text(child)
            if inner:
                texts.append(inner)

    return "".join(texts)
