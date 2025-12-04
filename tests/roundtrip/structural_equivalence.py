"""Structural equivalence comparison for OOXML documents.

This module compares two docx files at the XML structural level to determine
if they are semantically equivalent. It distinguishes between:

1. Semantic differences - actual content changes (test failures)
2. Ordering differences - same content in different order (acceptable)
3. Non-semantic differences - cleanup of unused metadata (acceptable)

Usage:
    from tests.roundtrip.structural_equivalence import compare_documents, EquivalenceResult

    result = compare_documents(original_path, roundtrip_path)
    if result.is_equivalent:
        print("Documents are structurally equivalent")
    else:
        print(f"Semantic differences: {result.semantic_diffs}")
"""

import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from lxml import etree


# XML containers where child element order is semantically irrelevant
UNORDERED_CONTAINERS = {
    "{http://schemas.openxmlformats.org/package/2006/content-types}Types",
    "{http://schemas.openxmlformats.org/package/2006/relationships}Relationships",
}


@dataclass
class EquivalenceResult:
    """Result of structural equivalence comparison."""

    is_equivalent: bool
    """True if documents are semantically equivalent (no content loss)."""

    is_identical: bool
    """True if documents are exactly identical (no differences at all)."""

    semantic_diffs: Dict[str, List[str]] = field(default_factory=dict)
    """Semantic differences by part name - these are test failures."""

    ordering_diffs: Dict[str, List[str]] = field(default_factory=dict)
    """Ordering differences - acceptable, content is the same."""

    non_semantic: List[str] = field(default_factory=list)
    """Non-semantic changes like cleanup of unused metadata."""

    file_diffs: Dict[str, List[str]] = field(default_factory=dict)
    """File-level differences (missing/added files)."""

    def __str__(self) -> str:
        if self.is_identical:
            return "IDENTICAL"
        elif self.is_equivalent:
            parts = ["EQUIVALENT"]
            if self.ordering_diffs:
                parts.append(f"({len(self.ordering_diffs)} ordering changes)")
            if self.non_semantic:
                parts.append(f"({len(self.non_semantic)} non-semantic changes)")
            return " ".join(parts)
        else:
            return f"DIFFERS ({len(self.semantic_diffs)} parts with semantic differences)"


def _is_directory_entry(name: str, zipf: zipfile.ZipFile) -> bool:
    """Check if a ZIP entry is just a directory marker."""
    return name.endswith("/") or zipf.getinfo(name).file_size == 0


def _is_empty_rels(content: bytes) -> bool:
    """Check if a .rels file is effectively empty (no relationships)."""
    try:
        root = etree.fromstring(content)
        return len(root) == 0
    except Exception:
        return False


def _get_used_extensions(zipf: zipfile.ZipFile) -> Set[str]:
    """Get file extensions actually used in the ZIP."""
    extensions = set()
    for name in zipf.namelist():
        if "." in name and not name.endswith("/"):
            ext = name.rsplit(".", 1)[1].lower()
            extensions.add(ext)
    return extensions


def _is_unused_content_type_default(
    entry: Tuple[str, Tuple], used_extensions: Set[str]
) -> bool:
    """Check if a Default content type entry is for an unused extension."""
    tag, attrs = entry
    if tag != "Default":
        return False
    attrs_dict = dict(attrs)
    ext = attrs_dict.get("Extension", "").lower()
    return ext not in used_extensions


def _is_orphan_content_type_override(
    entry: Tuple[str, Tuple], existing_parts: Set[str]
) -> bool:
    """Check if an Override content type entry refers to a non-existent part."""
    tag, attrs = entry
    if tag != "Override":
        return False
    attrs_dict = dict(attrs)
    part_name = attrs_dict.get("PartName", "").lstrip("/")
    return part_name not in existing_parts


def _is_element(node) -> bool:
    """Check if a node is an Element (not Comment, PI, etc.)."""
    # Element nodes have string tags; Comment/PI have callable tags
    return isinstance(node.tag, str)


def _element_children(elem: etree._Element) -> list:
    """Get only Element children (filter out Comments, PIs, etc.)."""
    return [c for c in elem if _is_element(c)]


def _element_signature(elem: etree._Element) -> tuple:
    """Create a hashable signature for an element (for set comparison)."""
    attrs = tuple(sorted(elem.attrib.items()))
    text = (elem.text or "").strip()
    child_sigs = tuple(_element_signature(c) for c in _element_children(elem))
    return (elem.tag, attrs, text, child_sigs)


def _normalize_whitespace(text: Optional[str]) -> str:
    """Normalize whitespace for comparison."""
    if text is None:
        return ""
    return text.strip()


def _elements_equal(
    e1: etree._Element, e2: etree._Element, path: str = ""
) -> Tuple[List[str], List[str]]:
    """
    Recursively compare two XML elements for structural equality.

    Returns:
        Tuple of (semantic_diffs, ordering_diffs)
    """
    semantic_diffs = []
    ordering_diffs = []

    # Compare tags
    if e1.tag != e2.tag:
        semantic_diffs.append(f"{path}: tag differs: {e1.tag} vs {e2.tag}")
        return semantic_diffs, ordering_diffs

    # Build readable path
    local_tag = e1.tag.split("}")[-1] if "}" in e1.tag else e1.tag
    current_path = f"{path}/{local_tag}"
    is_unordered = e1.tag in UNORDERED_CONTAINERS

    # Compare attributes
    attrs1 = dict(e1.attrib)
    attrs2 = dict(e2.attrib)
    keys1 = set(attrs1.keys())
    keys2 = set(attrs2.keys())

    for key in keys1 - keys2:
        semantic_diffs.append(f"{current_path}: attr missing: {key}")
    for key in keys2 - keys1:
        semantic_diffs.append(f"{current_path}: attr added: {key}")
    for key in keys1 & keys2:
        if attrs1[key] != attrs2[key]:
            v1, v2 = attrs1[key][:50], attrs2[key][:50]
            semantic_diffs.append(f"{current_path}: attr {key}: '{v1}' -> '{v2}'")

    # Compare text
    text1 = _normalize_whitespace(e1.text)
    text2 = _normalize_whitespace(e2.text)
    if text1 != text2:
        semantic_diffs.append(f"{current_path}: text differs")

    # Compare tail
    tail1 = _normalize_whitespace(e1.tail)
    tail2 = _normalize_whitespace(e2.tail)
    if tail1 != tail2:
        semantic_diffs.append(f"{current_path}: tail differs")

    # Compare children (only Element nodes, not Comments/PIs)
    children1 = _element_children(e1)
    children2 = _element_children(e2)

    if is_unordered:
        # Set-based comparison for unordered containers
        sigs1 = Counter(_element_signature(c) for c in children1)
        sigs2 = Counter(_element_signature(c) for c in children2)

        if sigs1 != sigs2:
            missing = sigs1 - sigs2
            added = sigs2 - sigs1
            if missing:
                semantic_diffs.append(
                    f"{current_path}: {sum(missing.values())} children missing"
                )
            if added:
                semantic_diffs.append(
                    f"{current_path}: {sum(added.values())} children added"
                )
        elif [_element_signature(c) for c in children1] != [
            _element_signature(c) for c in children2
        ]:
            ordering_diffs.append(f"{current_path}: reordered")
    else:
        # Ordered comparison
        if len(children1) != len(children2):
            semantic_diffs.append(
                f"{current_path}: child count: {len(children1)} -> {len(children2)}"
            )

        for i, (c1, c2) in enumerate(zip(children1, children2)):
            sem, ord_ = _elements_equal(c1, c2, f"{current_path}[{i}]")
            semantic_diffs.extend(sem)
            ordering_diffs.extend(ord_)

    return semantic_diffs, ordering_diffs


def _compare_content_types(
    z1: zipfile.ZipFile, z2: zipfile.ZipFile
) -> Tuple[List[str], List[str], List[str], bool]:
    """
    Compare [Content_Types].xml accounting for unused defaults and orphan overrides.

    Returns:
        Tuple of (semantic_missing, semantic_added, non_semantic, reordered)
    """
    content1 = z1.read("[Content_Types].xml")
    content2 = z2.read("[Content_Types].xml")

    root1 = etree.fromstring(content1)
    root2 = etree.fromstring(content2)

    used_ext1 = _get_used_extensions(z1)
    used_ext2 = _get_used_extensions(z2)

    # Get actual parts in each archive (for orphan detection)
    parts1 = {f for f in z1.namelist() if not f.endswith("/")}
    parts2 = {f for f in z2.namelist() if not f.endswith("/")}

    def get_entries(root):
        entries = []
        for child in root:
            local_tag = child.tag.split("}")[-1]
            attrs = tuple(sorted(child.attrib.items()))
            entries.append((local_tag, attrs))
        return entries

    entries1 = set(get_entries(root1))
    entries2 = set(get_entries(root2))

    missing = entries1 - entries2
    added = entries2 - entries1

    # Filter out unused content type defaults and orphan overrides
    non_semantic = []
    semantic_missing = []

    for e in missing:
        if _is_unused_content_type_default(e, used_ext1):
            non_semantic.append(f"unused default: {dict(e[1]).get('Extension', 'unknown')}")
        elif _is_orphan_content_type_override(e, parts1):
            non_semantic.append(f"orphan override: {dict(e[1]).get('PartName', 'unknown')}")
        else:
            semantic_missing.append(e)

    semantic_added = [
        e for e in added if not _is_unused_content_type_default(e, used_ext2)
    ]

    # Check ordering (only if semantically equal)
    reordered = False
    if not semantic_missing and not semantic_added:
        entries1_list = get_entries(root1)
        entries2_list = get_entries(root2)
        entries1_used = [
            e for e in entries1_list if not _is_unused_content_type_default(e, used_ext1)
        ]
        entries2_used = [
            e for e in entries2_list if not _is_unused_content_type_default(e, used_ext2)
        ]
        reordered = entries1_used != entries2_used

    return semantic_missing, semantic_added, non_semantic, reordered


def compare_documents(
    original_path: Path, roundtrip_path: Path
) -> EquivalenceResult:
    """
    Compare two docx files for structural equivalence.

    Args:
        original_path: Path to the original document
        roundtrip_path: Path to the round-tripped document

    Returns:
        EquivalenceResult with detailed comparison information
    """
    result = EquivalenceResult(is_equivalent=True, is_identical=True)

    with zipfile.ZipFile(original_path, "r") as z1, zipfile.ZipFile(
        roundtrip_path, "r"
    ) as z2:
        # Get content files (ignore directory entries)
        files1 = {f for f in z1.namelist() if not _is_directory_entry(f, z1)}
        files2 = {f for f in z2.namelist() if not _is_directory_entry(f, z2)}

        # Handle [Content_Types].xml specially
        sem_missing, sem_added, non_sem, reordered = _compare_content_types(z1, z2)

        if sem_missing:
            result.semantic_diffs["[Content_Types].xml"] = [
                f"Missing: {e}" for e in sem_missing
            ]
        if sem_added:
            result.semantic_diffs.setdefault("[Content_Types].xml", []).extend(
                [f"Added: {e}" for e in sem_added]
            )
        if non_sem:
            result.non_semantic.extend(
                [f"[Content_Types].xml: {ns}" for ns in non_sem]
            )
            result.is_identical = False
        if reordered:
            result.ordering_diffs["[Content_Types].xml"] = ["reordered"]
            result.is_identical = False

        # Check file differences
        missing = files1 - files2
        added = files2 - files1

        for f in list(missing):
            if f.endswith(".rels"):
                content = z1.read(f)
                if _is_empty_rels(content):
                    result.non_semantic.append(f"Empty rels removed: {f}")
                    result.is_identical = False
                    missing.remove(f)

        if missing:
            result.file_diffs["missing"] = list(missing)
        if added:
            result.file_diffs["added"] = list(added)

        # Compare common XML files (except [Content_Types].xml)
        common = files1 & files2
        xml_files = [
            f
            for f in common
            if (f.endswith(".xml") or f.endswith(".rels"))
            and f != "[Content_Types].xml"
        ]

        for filename in sorted(xml_files):
            try:
                root1 = etree.fromstring(z1.read(filename))
                root2 = etree.fromstring(z2.read(filename))
                sem, ord_ = _elements_equal(root1, root2)
                if sem:
                    result.semantic_diffs[filename] = sem
                if ord_:
                    result.ordering_diffs[filename] = ord_
                    result.is_identical = False
            except Exception as e:
                result.semantic_diffs[filename] = [f"Error: {e}"]

        # Compare binary files
        binary_files = [
            f for f in common if not f.endswith(".xml") and not f.endswith(".rels")
        ]
        for filename in binary_files:
            if z1.read(filename) != z2.read(filename):
                result.semantic_diffs[filename] = ["Binary content differs"]

    # Determine final status
    has_semantic = (
        bool(result.file_diffs.get("missing"))
        or bool(result.file_diffs.get("added"))
        or bool(result.semantic_diffs)
    )

    if has_semantic:
        result.is_equivalent = False
        result.is_identical = False

    return result


def test_roundtrip_equivalence(docx_path: Path, tmp_path: Path) -> EquivalenceResult:
    """
    Test that a document maintains structural equivalence through round-trip.

    Args:
        docx_path: Path to the document to test
        tmp_path: Temporary directory for the round-tripped file

    Returns:
        EquivalenceResult from the comparison
    """
    from docx import Document

    # Read and save
    doc = Document(str(docx_path))
    roundtrip_path = tmp_path / docx_path.name
    doc.save(str(roundtrip_path))

    # Compare
    return compare_documents(docx_path, roundtrip_path)
