"""Revisions (track changes) proxy objects."""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import TYPE_CHECKING, Iterator

if TYPE_CHECKING:
    from docx.oxml.revisions import CT_RunTrackChange


class RevisionType(Enum):
    """Types of revisions in a document."""
    INSERTION = "insertion"
    DELETION = "deletion"
    FORMAT_CHANGE = "format_change"


class Revisions:
    """Collection of revisions (track changes) in a document.

    Provides read-only access to insertions, deletions, and formatting changes
    that have been made to a document with revision tracking enabled.

    Example::

        # Iterate over all revisions
        for revision in document.revisions:
            print(f"{revision.revision_type.value}: '{revision.text}' by {revision.author}")

        # Get only insertions
        for ins in document.revisions.insertions:
            print(f"Added: {ins.text}")

        # Get only deletions
        for deletion in document.revisions.deletions:
            print(f"Deleted: {deletion.text}")

        # Filter by author
        john_changes = [r for r in document.revisions if r.author == "John Doe"]
    """

    def __init__(
        self,
        insertions: list[CT_RunTrackChange],
        deletions: list[CT_RunTrackChange],
    ) -> None:
        self._insertions = insertions
        self._deletions = deletions

    def __iter__(self) -> Iterator[Revision]:
        """Iterate over all revisions (insertions and deletions)."""
        for ins in self._insertions:
            yield Revision(ins, RevisionType.INSERTION)
        for deletion in self._deletions:
            yield Revision(deletion, RevisionType.DELETION)

    def __len__(self) -> int:
        """Return the total number of revisions."""
        return len(self._insertions) + len(self._deletions)

    @property
    def insertions(self) -> list[Revision]:
        """List of all insertion revisions."""
        return [Revision(ins, RevisionType.INSERTION) for ins in self._insertions]

    @property
    def deletions(self) -> list[Revision]:
        """List of all deletion revisions."""
        return [Revision(d, RevisionType.DELETION) for d in self._deletions]

    @property
    def authors(self) -> list[str]:
        """List of unique authors who made revisions, in no particular order."""
        authors_set: set[str] = set()
        for ins in self._insertions:
            if ins.revision_author:
                authors_set.add(ins.revision_author)
        for deletion in self._deletions:
            if deletion.revision_author:
                authors_set.add(deletion.revision_author)
        return list(authors_set)

    def by_author(self, author: str) -> list[Revision]:
        """Return all revisions made by the specified author."""
        return [r for r in self if r.author == author]

    def accept_all(self) -> int:
        """Accept all revisions in the document.

        Returns the number of revisions accepted.

        Example::

            count = document.revisions.accept_all()
            print(f"Accepted {count} revisions")
        """
        count = 0
        # Process in reverse order to avoid issues with removed elements
        for ins in reversed(self._insertions):
            ins.accept()
            count += 1
        for deletion in reversed(self._deletions):
            deletion.accept()
            count += 1
        return count

    def reject_all(self) -> int:
        """Reject all revisions in the document.

        Returns the number of revisions rejected.

        Example::

            count = document.revisions.reject_all()
            print(f"Rejected {count} revisions")
        """
        count = 0
        # Process in reverse order to avoid issues with removed elements
        for ins in reversed(self._insertions):
            ins.reject()
            count += 1
        for deletion in reversed(self._deletions):
            deletion.reject()
            count += 1
        return count


class Revision:
    """Proxy for a single revision (tracked change).

    Represents an insertion, deletion, or formatting change made to the document.

    Attributes:
        revision_type: The type of change (INSERTION, DELETION, FORMAT_CHANGE)
        text: The text that was inserted or deleted
        author: The person who made the change
        date: When the change was made
        revision_id: Unique identifier for this revision
    """

    def __init__(
        self,
        track_change_elm: CT_RunTrackChange,
        revision_type: RevisionType,
    ) -> None:
        self._element = track_change_elm
        self._revision_type = revision_type

    @property
    def revision_type(self) -> RevisionType:
        """The type of this revision (INSERTION, DELETION, or FORMAT_CHANGE)."""
        return self._revision_type

    @property
    def text(self) -> str:
        """The text content of this revision.

        For insertions, this is the text that was added.
        For deletions, this is the text that was removed.
        """
        return self._element.revision_text

    @property
    def author(self) -> str:
        """The name of the person who made this change."""
        return self._element.revision_author

    @property
    def date(self) -> datetime | None:
        """The date/time when this change was made, or None if not available."""
        return self._element.revision_date

    @property
    def revision_id(self) -> int:
        """The unique identifier for this revision within the document."""
        return self._element.revision_id

    @property
    def is_insertion(self) -> bool:
        """True if this revision is an insertion."""
        return self._revision_type == RevisionType.INSERTION

    @property
    def is_deletion(self) -> bool:
        """True if this revision is a deletion."""
        return self._revision_type == RevisionType.DELETION

    def __repr__(self) -> str:
        text_preview = self.text[:30] + "..." if len(self.text) > 30 else self.text
        return f"Revision({self.revision_type.value}, '{text_preview}', by={self.author})"

    def accept(self) -> None:
        """Accept this revision, applying the change and removing markup.

        For insertions: The inserted text becomes regular text (w:ins removed).
        For deletions: The deleted text stays deleted (w:del removed).

        After calling accept(), this Revision object becomes invalid and should
        not be used further.

        Example::

            # Accept a specific revision
            for revision in document.revisions:
                if revision.author == "John":
                    revision.accept()
        """
        self._element.accept()

    def reject(self) -> None:
        """Reject this revision, reverting the change and removing markup.

        For insertions: The inserted text is removed entirely (w:ins and content removed).
        For deletions: The deleted text is restored (w:delText converted to w:t).

        After calling reject(), this Revision object becomes invalid and should
        not be used further.

        Example::

            # Reject all deletions
            for revision in document.revisions.deletions:
                revision.reject()
        """
        self._element.reject()
