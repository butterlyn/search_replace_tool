# SEARCH_REPLACE_TOOL

## Description

Finds and replaces all occurrences of a word or phrase in one or multiple Microsoft Word documents, and creates output copies of each documents with changes.

For any issues, contact nicholas.butterly@aemo.com.au

## Usage

1. Place one or several word documents to be searched in the `1_insert_word_documents_here` folder (NOTE: Only `.docx` files are supported).
2. Insert search-replace word pairs in a CSV in the `2_insert_search-replace_pairs_here` folder. First row of the CSV is ignored (reserved for headers). Able to handle multiple CSV files.
3. Find outputted documents in the `3_output` folder.

### Known Issues

- Does not replace words found in headers or footers (but does work for words in text boxes or shapes)
- Can take a long time to run, especially for large documents or for many search-replace pairs. As a rule of thumb, the WEM Rules takes approximately 1 minute per search-replace pair.

## Requirements

- Windows Operating System
- Microsoft Word

- Windows Operating System
- Microsoft Word
- Python package manager (pip, conda, etc.)

## Installation

### Installation requirements

- Python 3.X
- Pip
- pyinstaller

### Installation instructions

```
python setup.py sdist
```

Alternatively, assuming you have python installed, run `setup.bat` in either powershell or command prompt. Terminal must be python enabled, check by running `python --version`.

After installation, the `dist` folder will contain the tool.

## Future Improvements

- [ ] Add support for live editing within the same file, so that the user can see the changes being made as they are being made

##### Fix known issues

- [ ] Fix issue with not replacing words found in headers or footers
- [ ] Improve performance
  - [ ] Add support for multithreading
  - [ ] Add support for multiprocessing

##### Add support for more file types

- [ ] Add support for `.rtf` files
- [ ] Add support for `.pdf` files (via converting to `.docx`)

##### Add feature for changing text formatting

- [ ] Add support for highlighting
- [ ] Add support for bolding, italicising, and underlining

##### Usability improvements

- [ ] Add GUI
- [ ] Ability to specify any filepath(s) for a docx or directory containing docx files, and specify the CSV filepath(s)
- [ ] Create COM plugin for Word
- [ ] Improve installation verbosity
  - [ ] Add a progress bar
  - [ ] Display more information in installation terminal during installation (via a INFO logger)
