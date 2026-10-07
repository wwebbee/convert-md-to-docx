# md2docx

> Конвертер Markdown → Word (.docx). Текст остаётся без изменений — только формат.

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)]()
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Tests](https://img.shields.io/badge/tests-25_pass-green.svg)]()

## Установка

### Вариант 1 — из репозитория (рекомендуется)

```powershell
# клонируйте репозиторий
git clone https://github.com/USERNAME/md2docx.git
cd md2docx

# создайте виртуальное окружение (если ещё нет)
python -m venv .venv

# установите зависимость
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

### Вариант 2 — через pip (локально)

```powershell
pip install -e .
# теперь доступна команда md2docx без пути к скрипту
```

### Вариант 3 — глобально

```powershell
pip install python-docx
```

## Использование

### Базовый запуск

```powershell
# один файл → .docx рядом с исходником
md2docx.cmd файл.md

# пакетно
md2docx.cmd doc1.md doc2.md doc3.md

# все .md в текущей папке
md2docx.cmd --all
```

### Явный выходной файл / папка

```powershell
# конкретное имя результата
md2docx.cmd файл.md -o результат.docx

# все файлы → в папку out
md2docx.cmd --all -o out

# конкретная папка
md2docx.cmd docs\*.md -o out
```

### Прямой вызов через Python

```powershell
.\.venv\Scripts\python.exe md2docx.py файл.md
.\.venv\Scripts\python.exe md2docx.py файл.md -o out.docx
.\.venv\Scripts\python.exe md2docx.py --all
```

### Справка по CLI

```
usage: md2docx [-h] [-a] [-o OUT] [-V] [paths ...]

Convert Markdown files to Word (.docx) keeping the text unchanged.

positional arguments:
  paths          .md files or glob patterns

options:
  -h, --help     show this help message and exit
  -a, --all      convert every .md found in the given folder (default: current folder)
  -o, --out OUT  output .docx file (single input) or output folder
  -V, --version  show program version
```

## Поддерживаемая разметка

| Markdown | Результат в Word |
|---|---|
| `# Заголовок` — `######` | Heading 1 — Heading 6 |
| `---` | Горизонтальная линия |
| `| A | B |` | `|---|---|` | `| 1 | 2 |` | Таблица «Table Grid», залитая шапка, `<br>` → перенос строки |
| `` `код` `` | Шрифт Consolas 9pt |
| `**жирный**` | Жирный текст |
| `- пункт` / `* пункт` / `+ пункт` | Маркированный список |
| `1. пункт` | Нумерованный список (перезапуск с 1 для каждой новой группы) |
| ` ``` ` блок | Моноширинный текст с серой заливкой, отступы сохранены |
| `> цитата` | Стиль Quote |
| `---` | Горизонтальный разделитель |

### Ограничения

| Разметка | Поведение |
|---|---|
| Ссылки `[текст](url)` | Передаются как текст |
| Изображения `![](url)` | Передаются как текст |
| Курсив `*текст*`, зачёркивание `~~текст~~` | Передаются как текст |
| Вложенные списки | Все уровни → один уровень |
| Заголовки `#######` (7+ хешей) | Передаются как текст |
| Содержимое внутри `` `код` `` | Маркап не обрабатывается (CommonMark-совместимо) |

## Выходные коды

| Код | Значение |
|---|---|
| `0` | Все файлы конвертированы |
| `1` | Нет файлов для конвертации / ошибка |
| `2` | python-docx не установлен |

## Примеры

### Сценарий: конвертировать все тест-кейсы в docs/

```powershell
md2docx.cmd docs\*.md -o out
# → out/test-case-A.docx, out/test-case-B.docx, ...
```

### Сценарий: конвертировать конкретный файл с явным именем

```powershell
md2docx.cmd "тест-кейсы.md" -o "Тестовые_сценарии.docx"
# → Тестовые_сценарии.docx (в текущей папке)
```

### Сценарий: пакетная конвертация с отчётом

```powershell
md2docx.cmd a.md b.md c.md
# [ok]   a.md  ->  a.docx
# [skip] empty file: b.md
# [skip] not found: c.md
# Converted: 1/3
```

## Архитектура

```
md2docx/
├── md2docx.py          ← конвертер (один файл)
├── md2docx.cmd         ← обёртка для Windows
├── pyproject.toml      ← pip-установка
├── requirements.txt    ← зависимость
├── tests/              ← 25 тестов, unittest
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md
└── LICENSE
```

**Поток:** `parse()` → `build_document()` → `doc.save()` — всё в одном файле, без скрытой логики.

### Как работает нумерованный список

Каждая новая группа `1.` создаёт новый `numId` с `startOverride=1`. Это гарантирует, что в Word нумерация всегда начинается с 1 — даже если в Markdown пропущены номера.

### Как обрабатываются таблицы

- Строка-разделитель `|:--|--:|` не попадает в результат.
- `<br>` в ячейках → `<w:br/>` (перенос строки).
- Рваные строки (меньше ячеек) дополняются пустыми.
- Шапка: `bold=True` + цвет заливки `#D9E2F3`.

## Запуск тестов

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
# Ran 25 tests in 0.28s
# OK
```

Тесты используют только stdlib (`unittest`) — никаких дополнительных зависимостей.

## Вклад

См. [CONTRIBUTING.md](CONTRIBUTING.md).

## Лицензия

MIT — см. [LICENSE](LICENSE).
