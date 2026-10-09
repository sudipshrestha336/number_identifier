# Number Verifier

A Tkinter desktop app for reviewing the numbers in a **Word (.docx)** or **PDF** document.

It finds every numerical figure, highlights it, and attaches a comment asking the reviewer to confirm the value is correct and not misstated. Figures you have already verified can be skipped with an **ignore list**, typed in or loaded from an **Excel** file. Each run saves a copy of the document (your original is never changed) and a **CSV report** showing what was matched and what was commented.

## Features

- **Word and PDF** support from the same window. Choose the document type, attach the file, and run.
- **Finds** whole numbers, decimals, thousands separators (`1,234,567.89`), currency (`$`, `€`, `£`, `₹`, `¥`, `USD`, `Rs`), percentages, units such as `million`, `bn` or `lakh`, years, and bracketed negatives like `(30,807,514.42)`.
- **Highlights** each figure (colour of your choice) and adds a comment with editable wording (`{value}` and `{kind}` placeholders).
- **Ignore list** from an Excel file (every cell on every sheet) and/or typed entries.
- **CSV report** with every figure in document order, the ignore entry it matched, its status, its location and some context. The figures that were commented are listed again at the end.
- **Diagnostics in the log:** commented figures that are almost equal to an ignore entry ("near misses"), ignore entries that matched nothing, and PDF pages with no selectable text.
- **Formatting preserved:** Word runs are split only at the edges of a number and keep their formatting. PDFs get annotations appended, so the original content is untouched.

## Installation

Requires Python 3.9 or newer.

```bash
pip install -r requirements.txt
python number_comments_gui.py
```

`pymupdf` is only loaded when you choose PDF, so Word-only use works without it.

On Linux, tkinter may need `sudo apt install python3-tk`. On macOS, use the Python from python.org, because the Python bundled with Xcode often ships an outdated Tk.

## Usage

1. Select **Word (.docx)** or **PDF (.pdf)**.
2. Click **Browse…** to choose the input file. The output name fills in as `<name>_commented.docx` or `.pdf`; change it if you like.
3. Optional: choose an **Excel file** of numbers to ignore, and/or type extra entries.
4. Adjust the options (author, highlight colour, comment text, skip section numbers, CSV report).
5. Click **Add comments**. When it finishes you get the commented copy, plus `<name>_commented_report.csv` if the CSV option is ticked.

## How the ignore list matches

Matching is on the figure as written in the document, not on its calculated value.

| Rule | Example |
|---|---|
| Commas, spaces and case are ignored | `1000` matches `1,000`; `45 Million` matches `45 million` |
| Sign and brackets are ignored | `-5`, `(5)` and `5` are treated as the same figure |
| Trailing zeros are ignored | `12.5`, `12.50` and `13,733,724.00` vs `13733724` all match |
| A bare number also matches versions with a symbol or unit | `12.5` skips `12.5%`, `$12.5` and `12.5 million` |
| Whole figure only | `20` does not skip `2024` |
| Wildcards `*` and `?` | `20*` skips 2000, 2024, 20.5; `19??` skips 1900 to 1999 |
| Different spellings are different figures | `0.2` does not match `20%`; `1.9 million` does not match `1,900,000` |

In the typed box, entries are separated by commas, semicolons or new lines. Write `2200`, not `2,200`, because a comma separates entries.

**Excel notes**

- Numbers are read from the saved cell value. Formula cells need to have been calculated and saved in Excel.
- Cells formatted as percentages are read as shown (`0.2` formatted as `0%` becomes `20%`).
- **Round Excel numbers to the decimals shown in the cell** (on by default) removes floating-point noise such as `600669766.87000012` from `SUM` totals. Turn it off if you need exact stored values.
- **First row is a header** skips row 1 on every sheet.

## Limitations

**Word**
- Headers, footers, footnotes and text boxes are skipped (Word does not allow comments there).
- Auto-numbered list numbers are not stored as text and are not flagged.
- Numbers written as words ("five million") are not detected.
- Runs inside hyperlinks, tracked insertions or smart tags are not read, so a figure split across them may be misread.

**PDF**
- Scanned PDFs need OCR first; pages without selectable text are reported in the log.
- Password-protected PDFs must be unlocked first.
- Figures split across two lines are not detected.
- Comment display depends on the viewer (Acrobat and Foxit show comment text fully).

**General**
- The tool deliberately over-flags: things like `COVID-19` or phone numbers also get comments.
- Because sign and trailing zeros are ignored when matching, a figure with a wrong sign or extra zeros that otherwise equals an ignore entry will be skipped.

## Building an executable

PyInstaller cannot cross-compile, so build on the system you are targeting, using the same Python that has the packages installed.

```bash
python -m pip install pyinstaller
python -m PyInstaller --onefile --windowed --collect-all pymupdf number_comments_gui.py
```

On macOS, leave out `--onefile` (the result is `dist/number_comments_gui.app`). Unsigned apps need right-click > Open the first time. Some antivirus programs flag PyInstaller executables; this is a known false positive.

## Privacy

Everything runs locally. No document or ignore list is uploaded anywhere.

## License

Add a license of your choice (for example MIT) as a `LICENSE` file.
