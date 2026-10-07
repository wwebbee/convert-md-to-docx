# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] — 2026-07-10

Initial release.

### Added

- **Markdown → .docx converter** (`md2docx.py`): headings H1–H6, tables, inline code, bold, bullet and numbered lists, fenced code blocks, blockquotes, horizontal rules.
- **`-o/--out` CLI option**: output a single file with an explicit name, or direct all results into a folder (created automatically).
- **`--all` mode**: convert every `.md` file in the working directory (or a given folder) in one run.
- **`md2docx.cmd`** batch wrapper: resolves the project-local virtual environment, so the converter runs with a single command on Windows.
- **Numbered list restart**: every new `1.` group in Markdown gets its own `numId` with `startOverride=1`, so Word always restarts numbering at 1.
- **`<br>` inside table cells**: converted to `<w:br/>` (in-cell line breaks), matching the original text layout.
- **Header row styling**: table headers are bold with a light fill.
- **Test suite**: 25 tests (stdlib `unittest`, no extra dependencies) covering inline parsing, block parsing, table building, code blocks, numbered lists, and CLI behaviour.
- **Project files**: `pyproject.toml` (pip-installable), `requirements.txt`, `LICENSE` (MIT), `CONTRIBUTING.md`, `.gitignore`.

[1.0.0]: https://github.com/USERNAME/md2docx/releases/tag/v1.0.0
