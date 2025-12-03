"""SVG image header parser."""

from __future__ import annotations

import re
from typing import IO

from docx.image.constants import MIME_TYPE
from docx.image.image import BaseImageHeader


class Svg(BaseImageHeader):
    """Image header parser for SVG images.

    SVG (Scalable Vector Graphics) is an XML-based vector image format.
    Since SVG is resolution-independent, dimensions are interpreted as pixels
    at 96 DPI (the CSS/SVG standard).
    """

    @property
    def content_type(self) -> str:
        """MIME content type for this image, unconditionally `image/svg+xml`."""
        return MIME_TYPE.SVG

    @property
    def default_ext(self) -> str:
        """Default filename extension, always 'svg' for SVG images."""
        return "svg"

    @classmethod
    def from_stream(cls, stream: IO[bytes]) -> Svg:
        """Return an |Svg| instance with properties parsed from image in `stream`."""
        parser = _SvgParser.parse(stream)
        return cls(parser.px_width, parser.px_height, parser.horz_dpi, parser.vert_dpi)


class _SvgParser:
    """Parses an SVG image stream to extract image properties.

    SVG dimensions can be specified in several ways:
    1. Explicit width/height attributes (may include units)
    2. viewBox attribute (defines coordinate system)
    3. Both (width/height override viewBox for display size)

    Units supported: px, pt, em, ex, in, cm, mm, pc, % (percentage)
    Default unit is px if not specified.
    """

    # Standard DPI for SVG/CSS
    SVG_DPI = 96

    # Unit conversion factors to pixels at 96 DPI
    UNIT_TO_PX = {
        "px": 1.0,
        "pt": 96.0 / 72.0,  # 1 pt = 1/72 inch
        "pc": 96.0 / 6.0,  # 1 pica = 1/6 inch
        "in": 96.0,  # 1 inch = 96 px
        "cm": 96.0 / 2.54,  # 1 cm = 1/2.54 inch
        "mm": 96.0 / 25.4,  # 1 mm = 1/25.4 inch
        "em": 16.0,  # Assume 16px default font size
        "ex": 8.0,  # Assume x-height is half of em
    }

    def __init__(self, px_width: int, px_height: int):
        self._px_width = px_width
        self._px_height = px_height

    @classmethod
    def parse(cls, stream: IO[bytes]) -> _SvgParser:
        """Return an |_SvgParser| instance with properties from SVG in `stream`."""
        stream.seek(0)
        # Read enough to get the root SVG element (typically in first few KB)
        content = stream.read(8192).decode("utf-8", errors="ignore")

        width, height = cls._extract_dimensions(content)
        return cls(width, height)

    @classmethod
    def _extract_dimensions(cls, content: str) -> tuple[int, int]:
        """Extract width and height from SVG content.

        Returns (width, height) in pixels. Defaults to (300, 150) if not found
        (HTML5 default for replaced elements).
        """
        # Find the opening <svg> tag
        svg_match = re.search(r"<svg\s[^>]*>", content, re.IGNORECASE | re.DOTALL)
        if not svg_match:
            return 300, 150

        svg_tag = svg_match.group(0)

        # Try to extract width and height attributes
        width_attr = cls._extract_attribute(svg_tag, "width")
        height_attr = cls._extract_attribute(svg_tag, "height")

        # Try to extract viewBox
        viewbox = cls._extract_attribute(svg_tag, "viewBox")
        vb_width, vb_height = cls._parse_viewbox(viewbox)

        # Width and height attributes take precedence over viewBox
        if width_attr and height_attr:
            width = cls._parse_length(width_attr, vb_width)
            height = cls._parse_length(height_attr, vb_height)
        elif viewbox:
            # Use viewBox dimensions
            width = vb_width
            height = vb_height
        else:
            # Default dimensions (HTML5 default for replaced elements)
            width = 300
            height = 150

        return max(1, int(round(width))), max(1, int(round(height)))

    @classmethod
    def _extract_attribute(cls, tag: str, attr_name: str) -> str | None:
        """Extract attribute value from an XML tag string."""
        # Match attribute="value" or attribute='value'
        pattern = rf'{attr_name}\s*=\s*["\']([^"\']*)["\']'
        match = re.search(pattern, tag, re.IGNORECASE)
        return match.group(1) if match else None

    @classmethod
    def _parse_viewbox(cls, viewbox: str | None) -> tuple[float, float]:
        """Parse viewBox attribute and return (width, height)."""
        if not viewbox:
            return 0.0, 0.0

        # viewBox format: "min-x min-y width height"
        parts = viewbox.split()
        if len(parts) >= 4:
            try:
                width = float(parts[2])
                height = float(parts[3])
                return width, height
            except ValueError:
                pass
        return 0.0, 0.0

    @classmethod
    def _parse_length(cls, value: str, reference: float = 0.0) -> float:
        """Parse a length value with optional unit.

        Args:
            value: Length string like "100", "100px", "2.5in"
            reference: Reference value for percentage calculations

        Returns:
            Length in pixels
        """
        if not value:
            return 0.0

        value = value.strip()

        # Check for percentage
        if value.endswith("%"):
            try:
                percent = float(value[:-1])
                return (percent / 100.0) * reference
            except ValueError:
                return 0.0

        # Extract numeric value and unit
        match = re.match(r"^([+-]?[\d.]+)\s*([a-z]*)$", value, re.IGNORECASE)
        if not match:
            return 0.0

        try:
            num = float(match.group(1))
            unit = match.group(2).lower() if match.group(2) else "px"

            # Convert to pixels
            conversion = cls.UNIT_TO_PX.get(unit, 1.0)
            return num * conversion
        except ValueError:
            return 0.0

    @property
    def px_width(self) -> int:
        """Width of the image in pixels."""
        return self._px_width

    @property
    def px_height(self) -> int:
        """Height of the image in pixels."""
        return self._px_height

    @property
    def horz_dpi(self) -> int:
        """Horizontal DPI. SVG uses 96 DPI (CSS standard)."""
        return self.SVG_DPI

    @property
    def vert_dpi(self) -> int:
        """Vertical DPI. SVG uses 96 DPI (CSS standard)."""
        return self.SVG_DPI
