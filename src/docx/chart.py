"""Proxy objects for DrawingML charts.

Charts in Word documents are complex objects consisting of a drawing anchor
in the document, a chart part, and often an embedded spreadsheet for data.
"""

# pyright: reportPrivateUsage=false

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, List

from lxml import etree

if TYPE_CHECKING:
    pass


class Charts:
    """Collection of charts in a document.

    Provides iteration and indexing over charts found in the document.
    """

    def __init__(self, chart_elements: List[etree._Element]):
        self._chart_elements = chart_elements

    def __iter__(self) -> Iterator[Chart]:
        """Iterate over all charts."""
        for elem in self._chart_elements:
            yield Chart(elem)

    def __len__(self) -> int:
        """Return the number of charts."""
        return len(self._chart_elements)

    def __getitem__(self, index: int) -> Chart:
        """Return Chart at `index`."""
        return Chart(self._chart_elements[index])

    def __bool__(self) -> bool:
        """Return True if there are any charts."""
        return len(self._chart_elements) > 0


class Chart:
    """Proxy for a single chart in the document.

    Provides read-only access to basic chart properties.
    """

    def __init__(self, element: etree._Element):
        self._element = element

    @property
    def chart_type(self) -> str | None:
        """Return the type of chart (e.g., 'barChart', 'lineChart', 'pieChart').

        Returns None if the chart type cannot be determined.
        """
        # Chart type is determined by the first child of c:plotArea
        # that is not c:layout, c:catAx, c:valAx, etc.
        chart_types = [
            "barChart", "bar3DChart",
            "lineChart", "line3DChart",
            "pieChart", "pie3DChart", "doughnutChart",
            "areaChart", "area3DChart",
            "scatterChart",
            "bubbleChart",
            "radarChart",
            "surfaceChart", "surface3DChart",
            "stockChart",
        ]

        # Find the chart type element
        for type_name in chart_types:
            ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
            elem = self._element.find(f".//{ns}{type_name}")
            if elem is not None:
                return type_name

        return None

    @property
    def has_title(self) -> bool:
        """Return True if the chart has a title element."""
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        title_elem = self._element.find(f".//{ns}title")
        return title_elem is not None

    @property
    def title(self) -> str | None:
        """Return the chart title text, or None if no title.

        Note: Chart titles can be complex with multiple text runs.
        This returns a simple concatenation of all title text.
        """
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

        title_elem = self._element.find(f".//{ns}title")
        if title_elem is None:
            return None

        # Find all text elements within the title
        texts: List[str] = []
        for t_elem in title_elem.iter(f"{a_ns}t"):
            if t_elem.text:
                texts.append(str(t_elem.text))

        return "".join(texts) if texts else None

    @property
    def has_legend(self) -> bool:
        """Return True if the chart has a legend."""
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        legend_elem = self._element.find(f".//{ns}legend")
        return legend_elem is not None

    @property
    def series_count(self) -> int:
        """Return the number of data series in the chart."""
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        ser_elems = self._element.findall(f".//{ns}ser")
        return len(ser_elems)

    @property
    def series(self) -> List[ChartSeries]:
        """Return list of data series in the chart."""
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        ser_elems = self._element.findall(f".//{ns}ser")
        return [ChartSeries(elem) for elem in ser_elems]

    @property
    def categories(self) -> List[str]:
        """Return the category labels (x-axis values) for the chart.

        For charts with a category axis, this returns the labels.
        Returns an empty list if no categories are found.
        """
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"

        # Categories are typically in the first series' c:cat element
        # Try c:strRef/c:strCache first, then c:numRef/c:numCache
        cat_elem = self._element.find(f".//{ns}cat")
        if cat_elem is None:
            return []

        return _extract_cache_values(cat_elem, ns)


class ChartSeries:
    """Proxy for a single data series in a chart.

    Provides access to series name and data values.
    """

    def __init__(self, element: etree._Element):
        self._element = element
        self._ns = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
        self._a_ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    @property
    def name(self) -> str | None:
        """Return the series name, or None if not set.

        The series name is used in the legend.
        """
        # Series name is in c:tx/c:strRef/c:strCache/c:pt/c:v
        # or c:tx/c:v for literal values
        tx_elem = self._element.find(f"{self._ns}tx")
        if tx_elem is None:
            return None

        # Try literal value first
        v_elem = tx_elem.find(f"{self._ns}v")
        if v_elem is not None and v_elem.text:
            return str(v_elem.text)

        # Try string reference cache
        str_cache = tx_elem.find(f".//{self._ns}strCache")
        if str_cache is not None:
            pt = str_cache.find(f"{self._ns}pt")
            if pt is not None:
                v = pt.find(f"{self._ns}v")
                if v is not None and v.text:
                    return str(v.text)

        return None

    @property
    def values(self) -> List[float | None]:
        """Return the numeric data values for this series.

        Returns a list of floats. None values indicate missing data points.
        """
        val_elem = self._element.find(f"{self._ns}val")
        if val_elem is None:
            # For some chart types, values might be in c:yVal
            val_elem = self._element.find(f"{self._ns}yVal")
        if val_elem is None:
            return []

        return _extract_numeric_values(val_elem, self._ns)

    @property
    def categories(self) -> List[str]:
        """Return the category values for this series.

        For most charts, categories are shared across series.
        This returns the categories specific to this series.
        """
        cat_elem = self._element.find(f"{self._ns}cat")
        if cat_elem is None:
            # For scatter/bubble charts, try c:xVal
            cat_elem = self._element.find(f"{self._ns}xVal")
        if cat_elem is None:
            return []

        return _extract_cache_values(cat_elem, self._ns)

    @property
    def index(self) -> int:
        """Return the zero-based index of this series."""
        idx_elem = self._element.find(f"{self._ns}idx")
        if idx_elem is not None:
            val = idx_elem.get("val")
            if val is not None:
                return int(val)
        return 0


def _extract_cache_values(parent_elem: etree._Element, ns: str) -> List[str]:
    """Extract string values from a cache element (strCache or numCache).

    Used for extracting category labels.
    """
    values: List[str] = []

    # Try string cache first
    str_cache = parent_elem.find(f".//{ns}strCache")
    if str_cache is not None:
        pts = str_cache.findall(f"{ns}pt")
        # Sort by index to ensure correct order
        pts_with_idx: List[tuple[int, str]] = []
        for pt in pts:
            idx = int(pt.get("idx", "0"))
            v = pt.find(f"{ns}v")
            text = str(v.text) if v is not None and v.text else ""
            pts_with_idx.append((idx, text))
        pts_with_idx.sort(key=lambda x: x[0])
        values = [val for _, val in pts_with_idx]
        return values

    # Try number cache (some charts use numbers as categories)
    num_cache = parent_elem.find(f".//{ns}numCache")
    if num_cache is not None:
        pts = num_cache.findall(f"{ns}pt")
        pts_with_idx: List[tuple[int, str]] = []
        for pt in pts:
            idx = int(pt.get("idx", "0"))
            v = pt.find(f"{ns}v")
            text = str(v.text) if v is not None and v.text else ""
            pts_with_idx.append((idx, text))
        pts_with_idx.sort(key=lambda x: x[0])
        values = [val for _, val in pts_with_idx]

    return values


def _extract_numeric_values(parent_elem: etree._Element, ns: str) -> List[float | None]:
    """Extract numeric values from a numCache element.

    Used for extracting data series values.
    """
    values: List[float | None] = []

    num_cache = parent_elem.find(f".//{ns}numCache")
    if num_cache is None:
        return values

    # Get point count to pre-size the list
    pt_count_elem = num_cache.find(f"{ns}ptCount")
    pt_count = int(pt_count_elem.get("val", "0")) if pt_count_elem is not None else 0

    if pt_count > 0:
        values = [None] * pt_count

    pts = num_cache.findall(f"{ns}pt")
    for pt in pts:
        idx = int(pt.get("idx", "0"))
        v = pt.find(f"{ns}v")
        if v is not None and v.text:
            try:
                val = float(v.text)
                if idx < len(values):
                    values[idx] = val
                else:
                    # Extend list if needed
                    while len(values) <= idx:
                        values.append(None)
                    values[idx] = val
            except ValueError:
                pass

    return values
