"""|DocumentPart| and closely related objects."""

from __future__ import annotations

from typing import IO, TYPE_CHECKING, List, Tuple, cast

from docx.document import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.comments import CommentsPart
from docx.parts.customxml import CustomXmlPart, CustomXmlPropertiesPart
from docx.parts.footnotes import EndnotesPart, FootnotesPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.story import StoryPart
from docx.parts.styles import StylesPart
from docx.parts.theme import ThemePart
from docx.shape import InlineShapes
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from lxml import etree

    from docx.comments import Comments
    from docx.customxml import CustomXmlParts
    from docx.enum.style import WD_STYLE_TYPE
    from docx.footnotes import Endnotes, Footnotes
    from docx.opc.coreprops import CoreProperties
    from docx.settings import Settings
    from docx.styles.style import BaseStyle
    from docx.theme import Theme


class DocumentPart(StoryPart):
    """Main document part of a WordprocessingML (WML) package, aka a .docx file.

    Acts as broker to other parts such as image, core properties, and style parts. It
    also acts as a convenient delegate when a mid-document object needs a service
    involving a remote ancestor. The `Parented.part` property inherited by many content
    objects provides access to this part object for that purpose.
    """

    def add_footer_part(self):
        """Return (footer_part, rId) pair for newly-created footer part."""
        footer_part = FooterPart.new(self.package)
        rId = self.relate_to(footer_part, RT.FOOTER)
        return footer_part, rId

    def add_header_part(self):
        """Return (header_part, rId) pair for newly-created header part."""
        header_part = HeaderPart.new(self.package)
        rId = self.relate_to(header_part, RT.HEADER)
        return header_part, rId

    @property
    def comments(self) -> Comments:
        """|Comments| object providing access to the comments added to this document."""
        return self._comments_part.comments

    @property
    def core_properties(self) -> CoreProperties:
        """A |CoreProperties| object providing read/write access to the core properties
        of this document."""
        return self.package.core_properties

    @property
    def document(self):
        """A |Document| object providing access to the content of this document."""
        return Document(self._element, self)

    def drop_header_part(self, rId: str) -> None:
        """Remove related header part identified by `rId`."""
        self.drop_rel(rId)

    @property
    def endnotes(self) -> Endnotes:
        """|Endnotes| object providing access to the endnotes in this document."""
        from docx.footnotes import Endnotes

        return Endnotes(self._endnotes_part)

    def footer_part(self, rId: str):
        """Return |FooterPart| related by `rId`."""
        return self.related_parts[rId]

    @property
    def footnotes(self) -> Footnotes:
        """|Footnotes| object providing access to the footnotes in this document."""
        from docx.footnotes import Footnotes

        return Footnotes(self._footnotes_part)

    def get_style(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle:
        """Return the style in this document matching `style_id`.

        Returns the default style for `style_type` if `style_id` is |None| or does not
        match a defined style of `style_type`.
        """
        return self.styles.get_by_id(style_id, style_type)

    def get_style_id(self, style_or_name, style_type):
        """Return the style_id (|str|) of the style of `style_type` matching
        `style_or_name`.

        Returns |None| if the style resolves to the default style for `style_type` or if
        `style_or_name` is itself |None|. Raises if `style_or_name` is a style of the
        wrong type or names a style not present in the document.
        """
        return self.styles.get_style_id(style_or_name, style_type)

    def header_part(self, rId: str):
        """Return |HeaderPart| related by `rId`."""
        return self.related_parts[rId]

    @lazyproperty
    def inline_shapes(self):
        """The |InlineShapes| instance containing the inline shapes in the document."""
        return InlineShapes(self._element.body, self)

    @lazyproperty
    def numbering_part(self) -> NumberingPart:
        """A |NumberingPart| object providing access to the numbering definitions for this document.

        Creates an empty numbering part if one is not present.
        """
        try:
            return cast(NumberingPart, self.part_related_by(RT.NUMBERING))
        except KeyError:
            numbering_part = NumberingPart.new()
            self.relate_to(numbering_part, RT.NUMBERING)
            return numbering_part

    def save(self, path_or_stream: str | IO[bytes]):
        """Save this document to `path_or_stream`, which can be either a path to a
        filesystem location (a string) or a file-like object."""
        self.package.save(path_or_stream)

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the settings in the settings part of
        this document."""
        return self._settings_part.settings

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in the styles part of this
        document."""
        return self._styles_part.styles

    @property
    def theme(self) -> Theme:
        """A |Theme| object providing access to the document's theme.

        The theme defines the color scheme, font scheme, and effects used in the document.
        Returns a Theme object even if no theme part exists (with None/empty values).
        """
        from docx.theme import Theme

        return Theme(self._theme_part)

    @property
    def _comments_part(self) -> CommentsPart:
        """A |CommentsPart| object providing access to the comments added to this document.

        Creates a default comments part if one is not present.
        """
        try:
            return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
        except KeyError:
            assert self.package is not None
            comments_part = CommentsPart.default(self.package)
            self.relate_to(comments_part, RT.COMMENTS)
            return comments_part

    @property
    def _settings_part(self) -> SettingsPart:
        """A |SettingsPart| object providing access to the document-level settings for
        this document.

        Creates a default settings part if one is not present.
        """
        try:
            return cast(SettingsPart, self.part_related_by(RT.SETTINGS))
        except KeyError:
            settings_part = SettingsPart.default(self.package)
            self.relate_to(settings_part, RT.SETTINGS)
            return settings_part

    @property
    def _endnotes_part(self) -> EndnotesPart | None:
        """Instance of |EndnotesPart| for this document, or None if not present.

        Unlike other parts, endnotes are read-only and we don't create them on demand.
        """
        try:
            return cast(EndnotesPart, self.part_related_by(RT.ENDNOTES))
        except KeyError:
            return None

    @property
    def _footnotes_part(self) -> FootnotesPart | None:
        """Instance of |FootnotesPart| for this document, or None if not present.

        For read access, returns None if not present. Use `_get_or_add_footnotes_part()`
        to create the part on demand for write access.
        """
        try:
            return cast(FootnotesPart, self.part_related_by(RT.FOOTNOTES))
        except KeyError:
            return None

    def _get_or_add_footnotes_part(self) -> FootnotesPart:
        """Return the |FootnotesPart| for this document, creating one if needed.

        Creates a new FootnotesPart with separator footnotes if one doesn't exist.
        """
        footnotes_part = self._footnotes_part
        if footnotes_part is not None:
            return footnotes_part

        # Create a new footnotes part
        package = self.package
        assert package is not None
        footnotes_part = FootnotesPart.default(package)
        self.relate_to(footnotes_part, RT.FOOTNOTES)
        return footnotes_part

    @property
    def _styles_part(self) -> StylesPart:
        """Instance of |StylesPart| for this document.

        Creates an empty styles part if one is not present.
        """
        try:
            return cast(StylesPart, self.part_related_by(RT.STYLES))
        except KeyError:
            package = self.package
            assert package is not None
            styles_part = StylesPart.default(package)
            self.relate_to(styles_part, RT.STYLES)
            return styles_part

    @property
    def _theme_part(self) -> ThemePart | None:
        """Instance of |ThemePart| for this document, or None if not present.

        Unlike styles, themes are read-only and we don't create them on demand.
        """
        try:
            return cast(ThemePart, self.part_related_by(RT.THEME))
        except KeyError:
            return None

    @property
    def custom_xml_parts(self) -> CustomXmlParts:
        """A |CustomXmlParts| collection of custom XML parts in this document.

        Custom XML parts contain arbitrary XML data that can be bound to content
        controls for data-driven document generation.
        """
        from docx.customxml import CustomXmlParts

        parts: List[Tuple[CustomXmlPart, CustomXmlPropertiesPart | None]] = []
        package = self.package
        if package is None:
            return CustomXmlParts(parts)

        # Find all custom XML parts by relationship type
        for rel in package.iter_rels():
            if rel.is_external:
                continue
            if rel.reltype == RT.CUSTOM_XML:
                xml_part = cast(CustomXmlPart, rel.target_part)
                # Try to find the associated properties part
                props_part: CustomXmlPropertiesPart | None = None
                try:
                    props_part = cast(
                        CustomXmlPropertiesPart,
                        xml_part.part_related_by(RT.CUSTOM_XML_PROPS),
                    )
                except (KeyError, AttributeError):
                    pass
                parts.append((xml_part, props_part))

        return CustomXmlParts(parts)

    def add_custom_xml(
        self, xml_element: "etree._Element"
    ) -> tuple[CustomXmlPart, CustomXmlPropertiesPart]:
        """Add a new custom XML part to the document.

        Args:
            xml_element: The root XML element for the custom XML content.

        Returns:
            A tuple of (CustomXmlPart, CustomXmlPropertiesPart).
        """
        # Determine the next item number by checking existing custom XML parts
        package = self.package
        assert package is not None

        item_num = 1
        existing_parts = list(self.custom_xml_parts)
        if existing_parts:
            item_num = len(existing_parts) + 1

        # Create the custom XML and properties parts
        xml_part, props_part = CustomXmlPart.new(package, xml_element, item_num)

        # Add relationship from package to custom XML part
        package.rels.get_or_add(RT.CUSTOM_XML, xml_part)

        return xml_part, props_part
