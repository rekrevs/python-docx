"""Test structural equivalence for existing test documents in the repository.

These tests verify that all existing .docx test files maintain structural
equivalence through a read-save-read cycle.
"""

from pathlib import Path

import pytest

from tests.roundtrip.structural_equivalence import test_roundtrip_equivalence

# Discover all test document paths
REPO_ROOT = Path(__file__).parent.parent.parent
FEATURES_TEST_FILES = REPO_ROOT / "features" / "steps" / "test_files"
UNIT_TEST_FILES = REPO_ROOT / "tests" / "test_files"


def _collect_docx_files(directory: Path) -> list:
    """Collect all .docx files from a directory."""
    if not directory.exists():
        return []
    return sorted(directory.glob("*.docx"))


# Collect test files
FEATURES_DOCX = _collect_docx_files(FEATURES_TEST_FILES)
UNIT_DOCX = _collect_docx_files(UNIT_TEST_FILES)
ALL_DOCX = FEATURES_DOCX + UNIT_DOCX


class DescribeExistingDocumentRoundTrip:
    """Tests for round-trip structural equivalence of existing documents."""

    @pytest.mark.parametrize(
        "docx_path",
        FEATURES_DOCX,
        ids=[p.name for p in FEATURES_DOCX],
    )
    def it_preserves_features_test_files(self, docx_path: Path, tmp_path: Path):
        """Test that features/steps/test_files/*.docx files preserve structure."""
        result = test_roundtrip_equivalence(docx_path, tmp_path)

        assert result.is_equivalent, (
            f"Document {docx_path.name} has semantic differences:\n"
            f"  File diffs: {result.file_diffs}\n"
            f"  Semantic diffs: {result.semantic_diffs}"
        )

    @pytest.mark.parametrize(
        "docx_path",
        UNIT_DOCX,
        ids=[p.name for p in UNIT_DOCX],
    )
    def it_preserves_unit_test_files(self, docx_path: Path, tmp_path: Path):
        """Test that tests/test_files/*.docx files preserve structure."""
        result = test_roundtrip_equivalence(docx_path, tmp_path)

        assert result.is_equivalent, (
            f"Document {docx_path.name} has semantic differences:\n"
            f"  File diffs: {result.file_diffs}\n"
            f"  Semantic diffs: {result.semantic_diffs}"
        )


class DescribeEquivalenceDetails:
    """Tests that provide detailed equivalence information."""

    def it_reports_all_test_files_found(self):
        """Verify we found test files to test."""
        assert len(FEATURES_DOCX) > 0, "No features test files found"
        assert len(UNIT_DOCX) > 0, "No unit test files found"

    def it_can_test_all_documents_in_batch(self, tmp_path: Path):
        """Run all documents and report summary statistics."""
        results = {}
        for docx_path in ALL_DOCX:
            try:
                result = test_roundtrip_equivalence(docx_path, tmp_path)
                results[docx_path.name] = result
            except Exception as e:
                results[docx_path.name] = f"ERROR: {e}"

        # Count results
        identical = sum(
            1 for r in results.values() if isinstance(r, object) and hasattr(r, 'is_identical') and r.is_identical
        )
        equivalent = sum(
            1 for r in results.values() if isinstance(r, object) and hasattr(r, 'is_equivalent') and r.is_equivalent and not r.is_identical
        )
        differs = sum(
            1 for r in results.values() if isinstance(r, object) and hasattr(r, 'is_equivalent') and not r.is_equivalent
        )
        errors = sum(1 for r in results.values() if isinstance(r, str))

        # This test always passes but prints summary
        print(f"\n\nRound-trip Summary: {len(ALL_DOCX)} documents")
        print(f"  IDENTICAL:  {identical}")
        print(f"  EQUIVALENT: {equivalent}")
        print(f"  DIFFERS:    {differs}")
        print(f"  ERRORS:     {errors}")

        # Fail if any have semantic differences
        failed = [name for name, r in results.items()
                  if (isinstance(r, object) and hasattr(r, 'is_equivalent') and not r.is_equivalent)
                  or isinstance(r, str)]

        assert not failed, f"Documents with issues: {failed}"
