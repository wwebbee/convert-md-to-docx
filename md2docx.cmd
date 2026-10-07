@echo off
rem md2docx wrapper - converts Markdown to Word using the bundled venv.
rem Usage:  md2docx file.md
rem         md2docx a.md b.md
rem         md2docx docs\*.md
rem         md2docx --all
rem         md2docx file.md -o result.docx
setlocal
set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" "%~dp0md2docx.py" %*
endlocal
