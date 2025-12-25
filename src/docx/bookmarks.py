"""Bookmarks proxy objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from docx.oxml.ns import qn

if TYPE_CHECKING:
    from lxml.etree import _Element

    from docx.oxml.bookmarks import CT_Bookmark


class Bookmarks:
    """Collection of bookmarks in a document.

    Provides access to all bookmarks defined in the document body.
    Bookmarks are named ranges that can be used as targets for
    cross-references, hyperlinks, or programmatic document navigation.

    Example::

        # Iterate over all bookmarks
        for bookmark in document.bookmarks:
            print(f"{bookmark.name}: id={bookmark.bookmark_id}")

        # Get bookmark by name
        toc_bookmark = document.bookmarks.get("_Toc123456789")
        if toc_bookmark:
            print(f"Found TOC bookmark: {toc_bookmark.name}")

        # Get bookmark by index
        first_bookmark = document.bookmarks[0]

        # Check if bookmark exists
        if "MyBookmark" in document.bookmarks:
            print("Bookmark exists")
    """

    def __init__(self, bookmark_starts: list[CT_Bookmark]) -> None:
        self._bookmark_starts = bookmark_starts
        # Build name lookup for efficient access
        self._by_name: dict[str, CT_Bookmark] = {
            bm.bookmark_name: bm for bm in bookmark_starts
        }

    def __iter__(self) -> Iterator[Bookmark]:
        """Iterate over all bookmarks in document order."""
        for bm_start in self._bookmark_starts:
            yield Bookmark(bm_start)

    def __len__(self) -> int:
        """Return the number of bookmarks."""
        return len(self._bookmark_starts)

    def __getitem__(self, key: int | str) -> Bookmark | None:
        """Get a bookmark by index or name.

        Args:
            key: Either an integer index or a bookmark name string.

        Returns:
            The Bookmark if found, None otherwise.
        """
        if isinstance(key, int):
            if 0 <= key < len(self._bookmark_starts):
                return Bookmark(self._bookmark_starts[key])
            return None
        else:
            bm_start = self._by_name.get(key)
            return Bookmark(bm_start) if bm_start is not None else None

    def __contains__(self, name: str) -> bool:
        """Check if a bookmark with the given name exists."""
        return name in self._by_name

    def get(self, name: str) -> Bookmark | None:
        """Get a bookmark by name, or None if not found."""
        bm_start = self._by_name.get(name)
        return Bookmark(bm_start) if bm_start is not None else None

    @property
    def names(self) -> list[str]:
        """List of all bookmark names in document order."""
        return [bm.bookmark_name for bm in self._bookmark_starts]


class Bookmark:
    """Proxy for a single bookmark.

    A bookmark is a named range in the document. The actual range spans
    from the `w:bookmarkStart` element to the corresponding `w:bookmarkEnd`
    element with the same id.

    Attributes:
        name: The bookmark name (e.g., "_Toc123", "MyBookmark")
        bookmark_id: The numeric id linking start and end elements

    Common bookmark name prefixes:
        - `_Toc*`: Table of contents entries (auto-generated)
        - `_Ref*`: Cross-reference targets (auto-generated)
        - `_Hlk*`: Hyperlink targets (auto-generated)
        - `_GoBack`: Last editing position
        - Custom names for user-defined bookmarks
    """

    def __init__(self, bookmark_start: CT_Bookmark) -> None:
        self._bookmark_start = bookmark_start

    @property
    def name(self) -> str:
        """The name of this bookmark."""
        return self._bookmark_start.bookmark_name

    @property
    def bookmark_id(self) -> int:
        """The numeric id of this bookmark.

        This id is used internally to match bookmarkStart with bookmarkEnd.
        """
        return self._bookmark_start.bookmark_id

    @property
    def is_toc_entry(self) -> bool:
        """True if this is a table of contents bookmark."""
        return self.name.startswith("_Toc")

    @property
    def is_reference(self) -> bool:
        """True if this is a cross-reference bookmark."""
        return self.name.startswith("_Ref")

    @property
    def is_hyperlink(self) -> bool:
        """True if this is a hyperlink bookmark."""
        return self.name.startswith("_Hlk")

    @property
    def is_system(self) -> bool:
        """True if this is a system/auto-generated bookmark.

        System bookmarks have names starting with underscore.
        """
        return self.name.startswith("_")

    @property
    def bookmark_end(self) -> _Element | None:
        """The corresponding bookmarkEnd element, or None if not found.

        Searches in the parent and following siblings for a bookmarkEnd
        with a matching id.
        """
        bookmark_id = self.bookmark_id
        parent = self._bookmark_start.getparent()
        if parent is None:
            return None

        # Search for bookmarkEnd with matching id
        # It could be in the same parent or in ancestor's descendants
        root = parent
        while root.getparent() is not None:
            root = root.getparent()

        for end_elem in root.iter(qn("w:bookmarkEnd")):
            if end_elem.get(qn("w:id")) == str(bookmark_id):
                return end_elem
        return None

    @name.setter
    def name(self, value: str) -> None:
        """Set the name of this bookmark.

        Note: This only changes the bookmarkStart element's name attribute.
        """
        self._bookmark_start.set(qn("w:name"), value)

    def delete(self) -> None:
        """Delete this bookmark from the document.

        Removes both the bookmarkStart and corresponding bookmarkEnd elements.
        Content between them is preserved.
        """
        # Remove bookmarkEnd first
        end_elem = self.bookmark_end
        if end_elem is not None:
            end_parent = end_elem.getparent()
            if end_parent is not None:
                end_parent.remove(end_elem)

        # Remove bookmarkStart
        start_parent = self._bookmark_start.getparent()
        if start_parent is not None:
            start_parent.remove(self._bookmark_start)

    def __repr__(self) -> str:
        return f"Bookmark(name='{self.name}', id={self.bookmark_id})"
