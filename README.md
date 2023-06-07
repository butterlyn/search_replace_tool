# SEARCH_REPLACE_TOOL

## Description

Finds and replaces all occurrences of a word or phrase in one or multiple Microsoft Word documents, and creates output copies of each documents with changes.

For any issues, contact nicholas.butterly@aemo.com.au

## Requirements

### EXE version

- Windows Operating System
- Microsoft Word
- Administrative privileges

### Python version

- Windows Operating System
- Microsoft Word
- Python 3.X
- Python package manager (pip, conda, etc.)

## Installation

Nil.

## Usage

1. Place one or several word documents to be searched in the `input` folder (NOTE: Only `.docx` files are supported).
2. Insert search-replace word pairs in a CSV in the `search_replace` folder. First row of the CSV is ignored (reserved for headers). Able to handle multiple CSV files.
3. Run EXE version `search_replace.exe` if you are on Windows and have admin privileges. Otherwise, run Python version `search_replace.py`:

```
pip install -r requirements.txt
python search_replace.py
```

4. Find outputted documents in the `output` folder.

## Known Issues

- Does not replace words found in headers or footers (but does work for words in text boxes or shapes)
- Can take a long time to run, especially for large documents or for many search-replace pairs. As a rule of thumb, the WEM Rules takes approximately 1 minute per search-replace pair.

## Future Improvements

- [ ] Fix known issues
  - [ ] Fix issue with not replacing words found in headers or footers
  - [ ] Improve performance
    - [ ] Add support for multithreading
    - [ ] Add support for multiprocessing
- [ ] Add support for live editing within the same file, so that the user can see the changes being made as they are being made
- [ ] Add support for more file types
  - [ ] Add support for `.rtf` files
  - [ ] Add support for `.pdf` files (via converting to `.docx`)
- [ ] Add feature for changing text formatting
  - [ ] Add support for highlighting
  - [ ] Add support for bolding, italicising, and underlining
- [ ] Create COM plugin for Word
