# python-docx (xtend fork)

> Python library for creating/reading/updating Microsoft Word .docx files, extended with OOXML support.

## Quick Start

```bash
uv sync                        # Install dependencies
source .venv/bin/activate      # Activate venv
pytest tests/ -q               # Run unit tests
behave features/               # Run acceptance tests
```

## Tech Stack

- Python 3.9+, lxml, typing_extensions
- Testing: pytest, behave (BDD)
- Quality: pyright (strict), ruff

## Architecture

```yaml
layers:  # top to bottom
  - api.py: Document() entry point
  - proxy: Document, Paragraph, Run, Table, Section, Style, Font
  - oxml: CT_Document, CT_P, CT_R, CT_Tbl (declarative via xmlchemy)
  - lxml: Python XML processing
  - opc: ZIP package, Parts, Relationships, ContentTypes

src/docx/:
  api.py: public entry point
  document.py: Document class
  oxml/: CT_* element classes (XML ↔ Python mapping)
  parts/: DocumentPart, StylesPart, NumberingPart, FootnotesPart
  opc/: Open Packaging Conventions (ZIP structure)
  text/: Paragraph, Run, Font, Hyperlink
  styles/: Style, LatentStyles
  shape.py: InlineShape, FloatingShape
  table.py: Table, Row, Cell
  section.py: Section, Header, Footer

namespaces:  # in oxml/ns.py
  core: w, r, wp, a, pic, m, c, dgm
  extensions: w14, w15, w16, wp14, a14, asvg
  compat: mc, wps, wpg, v, o

tests/: unit tests (pytest, it_*/its_* naming)
features/: BDD acceptance tests (behave/gherkin)
wotan/:
  backlog.json: task index
  dev-log/: task files (T-NNNN.md)
  docs/: feature analysis, specifications
```

## Commands

```bash
# Testing
pytest tests/ -q                    # Unit tests
behave features/                    # Acceptance tests
pytest tests/ -q && behave          # Full suite

# Quality
pyright                             # Type checking
ruff check                          # Linting
ruff check --fix                    # Auto-fix lint issues

## Test Style

- Pytest with descriptive names: `it_does_something`, `its_property_returns_expected`
- Behave for acceptance: Gherkin features in `features/`
- Never weaken or skip regression tests without justification

## Extension Features

This fork adds read/write support for:
- Content controls (SDT), Fields, Footnotes/Endnotes
- Bookmarks, Track changes, Floating images, Text boxes
- Themes, Math equations, Charts, SmartArt, Custom XML

See `wotan/docs/` for specifications and `README.md` for API examples.

## OOXML References

- ECMA-376 5th Edition (Office Open XML)
- MS-DOCX Word Extensions documentation
- Analysis docs in `wotan/docs/`

## Verification

Before completing work:
```bash
pytest tests/ -q && behave features/ && pyright && ruff check
```
