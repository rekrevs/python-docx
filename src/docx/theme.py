"""Theme and related proxy objects for document themes.

A theme defines the color scheme, font scheme, and effects for a document.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from docx.shared import RGBColor

if TYPE_CHECKING:
    from docx.oxml.theme import CT_Color, CT_ColorScheme, CT_FontScheme, CT_OfficeStyleSheet
    from docx.parts.theme import ThemePart


class Theme:
    """Proxy for the document theme (``<a:theme>`` element).

    Provides access to the theme's color scheme, font scheme, and name.

    Example::

        # Access theme
        theme = document.theme

        # Theme name (e.g., "Office Theme")
        print(theme.name)

        # Theme colors
        print(theme.colors.accent1)  # RGBColor(0x4F, 0x81, 0xBD)

        # Theme fonts
        print(theme.fonts.major_latin)  # "Calibri Light"
    """

    def __init__(self, theme_part: ThemePart | None):
        self._theme_part = theme_part

    @property
    def _theme_elm(self) -> CT_OfficeStyleSheet | None:
        """Return the ``<a:theme>`` element, or None if no theme part."""
        if self._theme_part is None:
            return None
        return self._theme_part.theme_element

    @property
    def name(self) -> str:
        """The name of this theme (e.g., 'Office Theme'), or empty string if not set."""
        theme_elm = self._theme_elm
        if theme_elm is None or theme_elm.name is None:
            return ""
        return theme_elm.name

    @property
    def colors(self) -> ThemeColors:
        """A |ThemeColors| object providing access to the 12 theme colors."""
        theme_elm = self._theme_elm
        clr_scheme = None
        if theme_elm is not None and theme_elm.themeElements is not None:
            clr_scheme = theme_elm.themeElements.clrScheme
        return ThemeColors(clr_scheme)

    @property
    def fonts(self) -> ThemeFonts:
        """A |ThemeFonts| object providing access to the theme fonts."""
        theme_elm = self._theme_elm
        font_scheme = None
        if theme_elm is not None and theme_elm.themeElements is not None:
            font_scheme = theme_elm.themeElements.fontScheme
        return ThemeFonts(font_scheme)


class ThemeColors:
    """Provides access to the 12 theme colors.

    Theme colors are:
    - dark1, light1: Primary text/background colors
    - dark2, light2: Secondary text/background colors
    - accent1-6: Six accent colors for emphasis
    - hyperlink: Color for hyperlinks
    - followed_hyperlink: Color for visited hyperlinks
    """

    def __init__(self, clr_scheme: CT_ColorScheme | None):
        self._clr_scheme = clr_scheme

    def _get_color(self, color_elm: CT_Color | None) -> RGBColor | None:
        """Convert a theme color element to an RGBColor."""
        if color_elm is None:
            return None
        rgb_str = color_elm.rgb_color
        if rgb_str is None or len(rgb_str) != 6:
            return None
        # Parse hex string to RGB using RGBColor.from_string
        try:
            return RGBColor.from_string(rgb_str)
        except (ValueError, TypeError):
            return None

    def _set_color(self, color_name: str, value: RGBColor | str | None) -> None:
        """Set a theme color by name.

        Args:
            color_name: The color element name (e.g., 'dk1', 'accent1').
            value: RGBColor or hex string (e.g., 'FF0000'), or None to clear.
        """
        if self._clr_scheme is None:
            raise ValueError("No color scheme available - cannot modify colors")

        # Get the color element (e.g., self._clr_scheme.dk1)
        color_elm: CT_Color | None = getattr(self._clr_scheme, color_name, None)
        if color_elm is None:
            # Create the color element if it doesn't exist
            add_method = getattr(self._clr_scheme, f"_add_{color_name}", None)
            if add_method is None:
                raise ValueError(f"Cannot create color element '{color_name}'")
            color_elm = add_method()

        # Convert RGBColor to hex string if needed
        if isinstance(value, RGBColor):
            hex_str = f"{value[0]:02X}{value[1]:02X}{value[2]:02X}"
        elif value is None:
            hex_str = "000000"  # Default to black
        else:
            hex_str = str(value).lstrip("#").upper()

        # color_elm is guaranteed to be a CT_Color here (either existed or was created)
        assert color_elm is not None
        color_elm.rgb_color = hex_str

    @property
    def dark1(self) -> RGBColor | None:
        """Dark 1 color, typically used for main text."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.dk1)

    @dark1.setter
    def dark1(self, value: RGBColor | str) -> None:
        self._set_color("dk1", value)

    @property
    def light1(self) -> RGBColor | None:
        """Light 1 color, typically used for background."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.lt1)

    @light1.setter
    def light1(self, value: RGBColor | str) -> None:
        self._set_color("lt1", value)

    @property
    def dark2(self) -> RGBColor | None:
        """Dark 2 color, typically used for secondary text."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.dk2)

    @dark2.setter
    def dark2(self, value: RGBColor | str) -> None:
        self._set_color("dk2", value)

    @property
    def light2(self) -> RGBColor | None:
        """Light 2 color, typically used for secondary background."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.lt2)

    @light2.setter
    def light2(self, value: RGBColor | str) -> None:
        self._set_color("lt2", value)

    @property
    def accent1(self) -> RGBColor | None:
        """Accent 1 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent1)

    @accent1.setter
    def accent1(self, value: RGBColor | str) -> None:
        self._set_color("accent1", value)

    @property
    def accent2(self) -> RGBColor | None:
        """Accent 2 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent2)

    @accent2.setter
    def accent2(self, value: RGBColor | str) -> None:
        self._set_color("accent2", value)

    @property
    def accent3(self) -> RGBColor | None:
        """Accent 3 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent3)

    @accent3.setter
    def accent3(self, value: RGBColor | str) -> None:
        self._set_color("accent3", value)

    @property
    def accent4(self) -> RGBColor | None:
        """Accent 4 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent4)

    @accent4.setter
    def accent4(self, value: RGBColor | str) -> None:
        self._set_color("accent4", value)

    @property
    def accent5(self) -> RGBColor | None:
        """Accent 5 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent5)

    @accent5.setter
    def accent5(self, value: RGBColor | str) -> None:
        self._set_color("accent5", value)

    @property
    def accent6(self) -> RGBColor | None:
        """Accent 6 color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.accent6)

    @accent6.setter
    def accent6(self, value: RGBColor | str) -> None:
        self._set_color("accent6", value)

    @property
    def hyperlink(self) -> RGBColor | None:
        """Hyperlink color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.hlink)

    @hyperlink.setter
    def hyperlink(self, value: RGBColor | str) -> None:
        self._set_color("hlink", value)

    @property
    def followed_hyperlink(self) -> RGBColor | None:
        """Followed (visited) hyperlink color."""
        if self._clr_scheme is None:
            return None
        return self._get_color(self._clr_scheme.folHlink)

    @followed_hyperlink.setter
    def followed_hyperlink(self, value: RGBColor | str) -> None:
        self._set_color("folHlink", value)

    @property
    def scheme_name(self) -> str:
        """The name of this color scheme (e.g., 'Office')."""
        if self._clr_scheme is None or self._clr_scheme.name is None:
            return ""
        return self._clr_scheme.name


class ThemeFonts:
    """Provides access to theme font definitions.

    Theme fonts are divided into:
    - Major fonts: Used for headings
    - Minor fonts: Used for body text

    Each category has fonts for:
    - Latin: Western text
    - East Asian: CJK text
    - Complex Script: RTL and Indic scripts
    """

    def __init__(self, font_scheme: CT_FontScheme | None):
        self._font_scheme = font_scheme

    def _set_font(self, collection: str, script: str, typeface: str) -> None:
        """Set a font typeface for a specific collection and script.

        Args:
            collection: 'majorFont' or 'minorFont'
            script: 'latin', 'ea', or 'cs'
            typeface: The font name to set
        """
        if self._font_scheme is None:
            raise ValueError("No font scheme available - cannot modify fonts")

        # Get or create the font collection (majorFont or minorFont)
        font_collection = getattr(self._font_scheme, collection, None)
        if font_collection is None:
            add_method = getattr(self._font_scheme, f"_add_{collection}", None)
            if add_method is None:
                raise ValueError(f"Cannot create font collection '{collection}'")
            font_collection = add_method()

        # Get or create the script element (latin, ea, or cs)
        script_elm = getattr(font_collection, script, None)
        if script_elm is None:
            add_method = getattr(font_collection, f"_add_{script}", None)
            if add_method is None:
                raise ValueError(f"Cannot create script element '{script}'")
            script_elm = add_method()

        script_elm.typeface = typeface

    @property
    def scheme_name(self) -> str:
        """The name of this font scheme (e.g., 'Office')."""
        if self._font_scheme is None or self._font_scheme.name is None:
            return ""
        return self._font_scheme.name

    @property
    def major_latin(self) -> str:
        """The major (heading) font for Latin text."""
        if self._font_scheme is None:
            return ""
        major = self._font_scheme.majorFont
        if major is None or major.latin is None:
            return ""
        return major.latin.typeface or ""

    @major_latin.setter
    def major_latin(self, value: str) -> None:
        self._set_font("majorFont", "latin", value)

    @property
    def major_east_asian(self) -> str:
        """The major (heading) font for East Asian text."""
        if self._font_scheme is None:
            return ""
        major = self._font_scheme.majorFont
        if major is None or major.ea is None:
            return ""
        return major.ea.typeface or ""

    @major_east_asian.setter
    def major_east_asian(self, value: str) -> None:
        self._set_font("majorFont", "ea", value)

    @property
    def major_complex_script(self) -> str:
        """The major (heading) font for Complex Script text."""
        if self._font_scheme is None:
            return ""
        major = self._font_scheme.majorFont
        if major is None or major.cs is None:
            return ""
        return major.cs.typeface or ""

    @major_complex_script.setter
    def major_complex_script(self, value: str) -> None:
        self._set_font("majorFont", "cs", value)

    @property
    def minor_latin(self) -> str:
        """The minor (body) font for Latin text."""
        if self._font_scheme is None:
            return ""
        minor = self._font_scheme.minorFont
        if minor is None or minor.latin is None:
            return ""
        return minor.latin.typeface or ""

    @minor_latin.setter
    def minor_latin(self, value: str) -> None:
        self._set_font("minorFont", "latin", value)

    @property
    def minor_east_asian(self) -> str:
        """The minor (body) font for East Asian text."""
        if self._font_scheme is None:
            return ""
        minor = self._font_scheme.minorFont
        if minor is None or minor.ea is None:
            return ""
        return minor.ea.typeface or ""

    @minor_east_asian.setter
    def minor_east_asian(self, value: str) -> None:
        self._set_font("minorFont", "ea", value)

    @property
    def minor_complex_script(self) -> str:
        """The minor (body) font for Complex Script text."""
        if self._font_scheme is None:
            return ""
        minor = self._font_scheme.minorFont
        if minor is None or minor.cs is None:
            return ""
        return minor.cs.typeface or ""

    @minor_complex_script.setter
    def minor_complex_script(self, value: str) -> None:
        self._set_font("minorFont", "cs", value)
