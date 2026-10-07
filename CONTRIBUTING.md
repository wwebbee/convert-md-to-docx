# Contributing

Спасибо за интерес к проекту! Этот документ описывает, как внести изменения.

## Before you start

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/short-description`.
3. Make your changes.
4. Run the tests: `python -m unittest discover -s tests`.
5. Commit with a clear message, then open a pull request.

## Code style

- Python 3.8+ — no `match` statements, no walrus operators in hot paths.
- Line length: up to 100 characters.
- `snake_case` for functions and variables, `PascalCase` for classes.
- Docstrings are concise (one line when obvious, longer only when non-obvious).
- Prefer the standard library; new dependencies must be discussed first.

## How the converter works

```
md2docx.py
  parse(lines)          →  list[tuple[str, Any]]   # (kind, payload)
  build_document(blocks) →  Document
  convert(src, dst)     →  None (writes the file)
```

Adding a new Markdown feature means:

1. Emit a new `(kind, payload)` in `parse()`.
2. Handle that kind in `build_document()`.
3. Add a test that covers both paths.

## Testing

Every change must keep the test suite green:

```powershell
python -m unittest discover -s tests -v
```

Tests use only `unittest` (stdlib). If your change touches inline parsing,
table building, or CLI behaviour, add or update the corresponding test case.

## Reporting bugs

Open an issue and include:

- The Markdown snippet that fails (minimal reproducible example).
- What you expected vs. what happened.
- Your OS and Python version.

## Feature requests

Open an issue, describe the use case. The project is intentionally small —
features that add heavy dependencies are unlikely to be merged.

## Commit messages

Use the imperative mood:

- `Add support for nested lists`
- `Fix numId collision in restart_numbering`
- `Update README with pip install instructions`

## License

By contributing, you agree that your contributions are licensed under the
[MIT License](LICENSE).
