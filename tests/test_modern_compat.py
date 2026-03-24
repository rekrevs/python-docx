"""Test round-trip fidelity of modern Office 365 / Copilot-era namespace extensions.

Validates that documents containing w16du:dateUtc (UTC revision timestamps) and
w16sfl:formattingAllowed (SDT format locking) attributes -- introduced in Word 2023+
and Word 2024+ respectively -- survive save/reload cycles without data loss.

Also verifies that these modern extensions coexist correctly with existing extension
features (SDT, fields, bookmarks, track changes) in a single document.
"""

from __future__ import annotations

from io import BytesIO

from docx import Document
from docx.oxml.ns import nsdecls, qn
from docx.oxml.parser import OxmlElement, parse_xml

# -- helpers -----------------------------------------------------------------


def _roundtrip(doc: Document) -> Document:
    """Save a Document to an in-memory buffer and reload it."""
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return Document(buf)


def _double_roundtrip(doc: Document) -> Document:
    """Save/reload a Document twice to verify stability."""
    return _roundtrip(_roundtrip(doc))


# -- w16du:dateUtc preservation ----------------------------------------------


class DescribeW16duDateUtcPreservation:
    """Verify that w16du:dateUtc attributes on revision marks survive round-trip."""

    def it_preserves_dateUtc_on_w_ins(self):
        """A w:ins element with w16du:dateUtc should retain the attribute after
        save/reload."""
        doc = Document()
        p = doc.add_paragraph("Before insertion")

        ins_xml = (
            "<w:ins %s %s "
            'w:id="1" w:author="TestUser" '
            'w:date="2025-06-15T10:30:00Z" '
            'w16du:dateUtc="2025-06-15T10:30:00Z">'
            "<w:r><w:t>inserted text</w:t></w:r>"
            "</w:ins>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        ins_elem = parse_xml(ins_xml)
        p._element.append(ins_elem)

        doc2 = _roundtrip(doc)

        ins_elements = doc2._element.body.xpath(".//w:ins")
        assert len(ins_elements) == 1

        ins_el = ins_elements[0]
        assert ins_el.get(qn("w:id")) == "1"
        assert ins_el.get(qn("w:author")) == "TestUser"
        assert ins_el.get(qn("w:date")) == "2025-06-15T10:30:00Z"
        assert ins_el.get(qn("w16du:dateUtc")) == "2025-06-15T10:30:00Z"

    def it_preserves_dateUtc_on_w_del(self):
        """A w:del element with w16du:dateUtc should retain the attribute."""
        doc = Document()
        p = doc.add_paragraph()

        del_xml = (
            "<w:del %s %s "
            'w:id="2" w:author="Reviewer" '
            'w:date="2025-07-20T14:45:00Z" '
            'w16du:dateUtc="2025-07-20T14:45:00Z">'
            "<w:r><w:delText>deleted text</w:delText></w:r>"
            "</w:del>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        del_elem = parse_xml(del_xml)
        p._element.append(del_elem)

        doc2 = _roundtrip(doc)

        del_elements = doc2._element.body.xpath(".//w:del")
        assert len(del_elements) == 1

        del_el = del_elements[0]
        assert del_el.get(qn("w:id")) == "2"
        assert del_el.get(qn("w:author")) == "Reviewer"
        assert del_el.get(qn("w16du:dateUtc")) == "2025-07-20T14:45:00Z"

    def it_preserves_dateUtc_on_w_rPrChange(self):
        """A w:rPrChange element with w16du:dateUtc should retain the attribute."""
        doc = Document()
        p = doc.add_paragraph()

        # A run with a formatting change tracked via rPrChange
        run_xml = (
            "<w:r %s %s>"
            "  <w:rPr>"
            "    <w:b/>"
            '    <w:rPrChange w:id="3" w:author="Editor" '
            '     w:date="2025-08-10T09:00:00Z" '
            '     w16du:dateUtc="2025-08-10T09:00:00Z">'
            "      <w:rPr/>"
            "    </w:rPrChange>"
            "  </w:rPr>"
            "  <w:t>formatted text</w:t>"
            "</w:r>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        run_elem = parse_xml(run_xml)
        p._element.append(run_elem)

        doc2 = _roundtrip(doc)

        rpr_changes = doc2._element.body.xpath(".//w:rPrChange")
        assert len(rpr_changes) == 1

        rpc = rpr_changes[0]
        assert rpc.get(qn("w:id")) == "3"
        assert rpc.get(qn("w:author")) == "Editor"
        assert rpc.get(qn("w16du:dateUtc")) == "2025-08-10T09:00:00Z"

    def it_preserves_dateUtc_on_w_pPrChange(self):
        """A w:pPrChange element with w16du:dateUtc should retain the attribute."""
        doc = Document()

        # A paragraph with a property change tracked via pPrChange
        p_xml = (
            "<w:p %s %s>"
            "  <w:pPr>"
            '    <w:jc w:val="center"/>'
            '    <w:pPrChange w:id="4" w:author="LayoutEditor" '
            '     w:date="2025-09-01T16:20:00Z" '
            '     w16du:dateUtc="2025-09-01T16:20:00Z">'
            "      <w:pPr/>"
            "    </w:pPrChange>"
            "  </w:pPr>"
            "  <w:r><w:t>centered paragraph</w:t></w:r>"
            "</w:p>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        p_elem = parse_xml(p_xml)
        doc._element.body.append(p_elem)

        doc2 = _roundtrip(doc)

        ppr_changes = doc2._element.body.xpath(".//w:pPrChange")
        assert len(ppr_changes) == 1

        ppc = ppr_changes[0]
        assert ppc.get(qn("w:id")) == "4"
        assert ppc.get(qn("w:author")) == "LayoutEditor"
        assert ppc.get(qn("w16du:dateUtc")) == "2025-09-01T16:20:00Z"

    def it_preserves_dateUtc_through_double_roundtrip(self):
        """w16du:dateUtc should survive two consecutive save/reload cycles."""
        doc = Document()
        p = doc.add_paragraph()

        ins_xml = (
            "<w:ins %s %s "
            'w:id="10" w:author="Copilot" '
            'w:date="2025-12-25T00:00:00Z" '
            'w16du:dateUtc="2025-12-25T00:00:00Z">'
            "<w:r><w:t>AI-generated text</w:t></w:r>"
            "</w:ins>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        p._element.append(parse_xml(ins_xml))

        doc3 = _double_roundtrip(doc)

        ins_elements = doc3._element.body.xpath(".//w:ins")
        assert len(ins_elements) == 1
        assert ins_elements[0].get(qn("w16du:dateUtc")) == "2025-12-25T00:00:00Z"

        # Also verify text content survived
        texts = doc3._element.body.xpath(".//w:ins/w:r/w:t/text()")
        assert texts == ["AI-generated text"]


# -- w16sfl:formattingAllowed preservation -----------------------------------


class DescribeW16sflFormattingAllowedPreservation:
    """Verify that w16sfl:formattingAllowed on SDT properties survives round-trip."""

    def it_preserves_formattingAllowed_on_block_sdt(self):
        """A block-level w:sdt with w16sfl:formattingAllowed='true' on its sdtPr
        should retain the attribute after save/reload."""
        doc = Document()

        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="modern-sdt"/>'
            '    <w:alias w:val="Modern Control"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>Content control text</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        sdt_elem = parse_xml(sdt_xml)
        doc._element.body.append(sdt_elem)

        doc2 = _roundtrip(doc)

        sdt_prs = doc2._element.body.xpath(".//w:sdtPr")
        assert len(sdt_prs) >= 1

        # Find the one with our tag
        matched = [
            pr
            for pr in sdt_prs
            if pr.xpath("./w:tag[@w:val='modern-sdt']")
        ]
        assert len(matched) == 1

        sdt_pr = matched[0]
        assert sdt_pr.get(qn("w16sfl:formattingAllowed")) == "true"

    def it_preserves_formattingAllowed_false_value(self):
        """w16sfl:formattingAllowed='false' should also be preserved (not optimized away)."""
        doc = Document()

        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="false">'
            '    <w:tag w:val="locked-sdt"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>Locked content</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)

        matched = [
            pr
            for pr in doc2._element.body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='locked-sdt']")
        ]
        assert len(matched) == 1
        assert matched[0].get(qn("w16sfl:formattingAllowed")) == "false"

    def it_preserves_formattingAllowed_with_existing_lock(self):
        """w16sfl:formattingAllowed should coexist with the traditional w:lock element."""
        doc = Document()

        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="lock-combo-sdt"/>'
            '    <w:lock w:val="sdtContentLocked"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>Locked body, formatting allowed</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)

        matched = [
            pr
            for pr in doc2._element.body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='lock-combo-sdt']")
        ]
        assert len(matched) == 1

        sdt_pr = matched[0]
        assert sdt_pr.get(qn("w16sfl:formattingAllowed")) == "true"

        # Also verify the traditional lock element survived
        lock_els = sdt_pr.xpath("./w:lock")
        assert len(lock_els) == 1
        assert lock_els[0].get(qn("w:val")) == "sdtContentLocked"

    def it_preserves_formattingAllowed_through_double_roundtrip(self):
        """w16sfl:formattingAllowed should survive two consecutive save/reload cycles."""
        doc = Document()

        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="double-rt-sdt"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>Double round-trip content</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc3 = _double_roundtrip(doc)

        matched = [
            pr
            for pr in doc3._element.body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='double-rt-sdt']")
        ]
        assert len(matched) == 1
        assert matched[0].get(qn("w16sfl:formattingAllowed")) == "true"


# -- Combined modern features in a single document --------------------------


class DescribeCombinedModernFeatures:
    """Verify that multiple modern extension features coexist in a single document."""

    def it_preserves_all_modern_attrs_in_one_document(self):
        """A document with both w16du:dateUtc on revisions and
        w16sfl:formattingAllowed on SDT should preserve everything."""
        doc = Document()

        # 1) Regular paragraph
        doc.add_paragraph("Regular paragraph before tracked changes.")

        # 2) Paragraph with tracked insertion bearing w16du:dateUtc
        p_ins = doc.add_paragraph()
        ins_xml = (
            "<w:ins %s %s "
            'w:id="100" w:author="Alice" '
            'w:date="2025-11-01T08:00:00Z" '
            'w16du:dateUtc="2025-11-01T08:00:00Z">'
            "<w:r><w:t>Alice inserted this.</w:t></w:r>"
            "</w:ins>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        p_ins._element.append(parse_xml(ins_xml))

        # 3) Paragraph with tracked deletion bearing w16du:dateUtc
        p_del = doc.add_paragraph()
        del_xml = (
            "<w:del %s %s "
            'w:id="101" w:author="Bob" '
            'w:date="2025-11-02T09:15:00Z" '
            'w16du:dateUtc="2025-11-02T09:15:00Z">'
            "<w:r><w:delText>Bob deleted this.</w:delText></w:r>"
            "</w:del>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        p_del._element.append(parse_xml(del_xml))

        # 4) SDT with w16sfl:formattingAllowed
        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="combined-sdt"/>'
            '    <w:alias w:val="Combined Test"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>SDT with modern locking</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        # 5) Regular paragraph after
        doc.add_paragraph("Regular paragraph after all features.")

        # Round-trip
        doc2 = _roundtrip(doc)
        body = doc2._element.body

        # Verify insertion
        ins_els = body.xpath(".//w:ins")
        assert len(ins_els) == 1
        assert ins_els[0].get(qn("w:id")) == "100"
        assert ins_els[0].get(qn("w16du:dateUtc")) == "2025-11-01T08:00:00Z"
        ins_texts = body.xpath(".//w:ins/w:r/w:t/text()")
        assert ins_texts == ["Alice inserted this."]

        # Verify deletion
        del_els = body.xpath(".//w:del")
        assert len(del_els) == 1
        assert del_els[0].get(qn("w:id")) == "101"
        assert del_els[0].get(qn("w16du:dateUtc")) == "2025-11-02T09:15:00Z"
        del_texts = body.xpath(".//w:del/w:r/w:delText/text()")
        assert del_texts == ["Bob deleted this."]

        # Verify SDT
        matched_prs = [
            pr
            for pr in body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='combined-sdt']")
        ]
        assert len(matched_prs) == 1
        assert matched_prs[0].get(qn("w16sfl:formattingAllowed")) == "true"

        # Verify regular paragraphs still present
        all_texts = [p.text for p in doc2.paragraphs]
        assert "Regular paragraph before tracked changes." in all_texts
        assert "Regular paragraph after all features." in all_texts

    def it_preserves_modern_attrs_alongside_bookmarks(self):
        """Modern namespace attributes should coexist with bookmarks."""
        doc = Document()
        p = doc.add_paragraph()

        # Bookmark start
        bm_start = OxmlElement("w:bookmarkStart")
        bm_start.set(qn("w:id"), "0")
        bm_start.set(qn("w:name"), "TestBookmark")
        p._element.insert(0, bm_start)

        # Tracked insertion with w16du:dateUtc
        ins_xml = (
            "<w:ins %s %s "
            'w:id="200" w:author="Carol" '
            'w:date="2025-10-15T12:00:00Z" '
            'w16du:dateUtc="2025-10-15T12:00:00Z">'
            "<w:r><w:t>bookmarked insertion</w:t></w:r>"
            "</w:ins>"
        ) % (nsdecls("w"), nsdecls("w16du"))
        p._element.append(parse_xml(ins_xml))

        # Bookmark end
        bm_end = OxmlElement("w:bookmarkEnd")
        bm_end.set(qn("w:id"), "0")
        p._element.append(bm_end)

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        # Verify bookmark
        bm_starts = body.xpath(".//w:bookmarkStart")
        assert len(bm_starts) >= 1
        assert bm_starts[0].get(qn("w:name")) == "TestBookmark"

        # Verify insertion with dateUtc
        ins_els = body.xpath(".//w:ins")
        assert len(ins_els) == 1
        assert ins_els[0].get(qn("w16du:dateUtc")) == "2025-10-15T12:00:00Z"

    def it_preserves_modern_attrs_alongside_fields(self):
        """Modern namespace attributes should coexist with field codes."""
        doc = Document()

        # Simple field
        p = doc.add_paragraph()
        fld_simple_xml = (
            '<w:fldSimple %s w:instr=" PAGE ">'
            "<w:r><w:t>1</w:t></w:r>"
            "</w:fldSimple>"
        ) % nsdecls("w")
        p._element.append(parse_xml(fld_simple_xml))

        # SDT with formattingAllowed
        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="field-combo-sdt"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>SDT near fields</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        # Verify field
        fld_simples = body.xpath(".//w:fldSimple")
        assert len(fld_simples) >= 1

        # Verify SDT
        matched = [
            pr
            for pr in body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='field-combo-sdt']")
        ]
        assert len(matched) == 1
        assert matched[0].get(qn("w16sfl:formattingAllowed")) == "true"

    def it_preserves_multiple_revisions_with_different_dateUtc_values(self):
        """Multiple tracked changes, each with distinct w16du:dateUtc values,
        should all be preserved independently."""
        doc = Document()

        timestamps = [
            ("300", "Alpha", "2025-03-01T10:00:00Z", "inserted by Alpha"),
            ("301", "Beta", "2025-03-02T11:00:00Z", "inserted by Beta"),
            ("302", "Gamma", "2025-03-03T12:00:00Z", "inserted by Gamma"),
        ]

        for rev_id, author, utc_date, text in timestamps:
            p = doc.add_paragraph()
            ins_xml = (
                "<w:ins %s %s "
                'w:id="%s" w:author="%s" '
                'w:date="%s" '
                'w16du:dateUtc="%s">'
                "<w:r><w:t>%s</w:t></w:r>"
                "</w:ins>"
            ) % (nsdecls("w"), nsdecls("w16du"), rev_id, author, utc_date, utc_date, text)
            p._element.append(parse_xml(ins_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        ins_els = body.xpath(".//w:ins")
        assert len(ins_els) == 3

        for i, (rev_id, author, utc_date, text) in enumerate(timestamps):
            el = ins_els[i]
            assert el.get(qn("w:id")) == rev_id
            assert el.get(qn("w:author")) == author
            assert el.get(qn("w16du:dateUtc")) == utc_date
            t_texts = el.xpath(".//w:t/text()")
            assert t_texts == [text]

    def it_preserves_multiple_sdts_with_formattingAllowed(self):
        """Multiple SDTs with w16sfl:formattingAllowed should each preserve their
        respective attribute values."""
        doc = Document()

        sdt_configs = [
            ("sdt-a", "true", "Content A"),
            ("sdt-b", "false", "Content B"),
            ("sdt-c", "true", "Content C"),
        ]

        for tag_val, fmt_val, text in sdt_configs:
            sdt_xml = (
                "<w:sdt %s %s>"
                '  <w:sdtPr w16sfl:formattingAllowed="%s">'
                '    <w:tag w:val="%s"/>'
                "  </w:sdtPr>"
                "  <w:sdtContent>"
                "    <w:p><w:r><w:t>%s</w:t></w:r></w:p>"
                "  </w:sdtContent>"
                "</w:sdt>"
            ) % (nsdecls("w"), nsdecls("w16sfl"), fmt_val, tag_val, text)
            doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        for tag_val, fmt_val, text in sdt_configs:
            matched = [
                pr
                for pr in body.xpath(".//w:sdtPr")
                if pr.xpath(f"./w:tag[@w:val='{tag_val}']")
            ]
            assert len(matched) == 1, f"SDT with tag '{tag_val}' not found"
            assert (
                matched[0].get(qn("w16sfl:formattingAllowed")) == fmt_val
            ), f"formattingAllowed mismatch for SDT '{tag_val}'"

    def it_preserves_w16sdtdh_alongside_w16sfl(self):
        """w16sdtdh:dataHash and w16sfl:formattingAllowed should coexist on the same
        sdtPr element without interfering."""
        doc = Document()

        sdt_xml = (
            "<w:sdt %s %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true"'
            '   w16sdtdh:dataHash="abc123hash">'
            '    <w:tag w:val="dual-ext-sdt"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>Dual extension SDT</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"), nsdecls("w16sdtdh"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        matched = [
            pr
            for pr in body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='dual-ext-sdt']")
        ]
        assert len(matched) == 1

        sdt_pr = matched[0]
        assert sdt_pr.get(qn("w16sfl:formattingAllowed")) == "true"
        assert sdt_pr.get(qn("w16sdtdh:dataHash")) == "abc123hash"


# -- Existing extension namespace coexistence --------------------------------


class DescribeExistingExtensionNamespacePreservation:
    """Verify that previously supported extension namespaces (w14, w15, w16se, etc.)
    continue to work correctly alongside the new w16du and w16sfl namespaces."""

    def it_preserves_w14_attributes_alongside_w16du(self):
        """w14 attributes on a run should coexist with w16du:dateUtc on a revision
        mark in the same paragraph."""
        doc = Document()
        p = doc.add_paragraph()

        # A tracked insertion with w16du:dateUtc that contains a run with a w14 attribute
        ins_xml = (
            "<w:ins %s %s %s "
            'w:id="400" w:author="ModernUser" '
            'w:date="2025-10-01T10:00:00Z" '
            'w16du:dateUtc="2025-10-01T10:00:00Z">'
            "  <w:r>"
            '    <w:rPr><w14:conflictIns w:id="500" w:author="ConflictUser"/></w:rPr>'
            "    <w:t>text with w14 and w16du</w:t>"
            "  </w:r>"
            "</w:ins>"
        ) % (nsdecls("w"), nsdecls("w14"), nsdecls("w16du"))
        p._element.append(parse_xml(ins_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        ins_els = body.xpath(".//w:ins")
        assert len(ins_els) == 1
        assert ins_els[0].get(qn("w16du:dateUtc")) == "2025-10-01T10:00:00Z"

        # Verify w14 element survived
        conflict_ins = body.xpath(".//w14:conflictIns")
        assert len(conflict_ins) == 1

    def it_preserves_w16cid_alongside_w16sfl(self):
        """w16cid:paraId attributes should coexist with w16sfl:formattingAllowed."""
        doc = Document()

        # A paragraph with w16cid:paraId followed by an SDT with w16sfl
        p_xml = (
            '<w:p %s %s w14:paraId="1A2B3C4D">'
            "  <w:r><w:t>paragraph with paraId</w:t></w:r>"
            "</w:p>"
        ) % (nsdecls("w"), nsdecls("w14"))
        doc._element.body.append(parse_xml(p_xml))

        sdt_xml = (
            "<w:sdt %s %s>"
            '  <w:sdtPr w16sfl:formattingAllowed="true">'
            '    <w:tag w:val="cid-combo-sdt"/>'
            "  </w:sdtPr>"
            "  <w:sdtContent>"
            "    <w:p><w:r><w:t>SDT near paraId paragraph</w:t></w:r></w:p>"
            "  </w:sdtContent>"
            "</w:sdt>"
        ) % (nsdecls("w"), nsdecls("w16sfl"))
        doc._element.body.append(parse_xml(sdt_xml))

        doc2 = _roundtrip(doc)
        body = doc2._element.body

        # Verify w14:paraId
        paras_with_id = body.xpath(".//w:p[@w14:paraId]")
        assert len(paras_with_id) >= 1
        assert paras_with_id[0].get(qn("w14:paraId")) == "1A2B3C4D"

        # Verify SDT with w16sfl
        matched = [
            pr
            for pr in body.xpath(".//w:sdtPr")
            if pr.xpath("./w:tag[@w:val='cid-combo-sdt']")
        ]
        assert len(matched) == 1
        assert matched[0].get(qn("w16sfl:formattingAllowed")) == "true"
