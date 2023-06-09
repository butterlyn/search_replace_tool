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
- Sometimes detects text to replace in shapes, but fails to replace them, giving an error message. May need to list words found in shapes and text boxes instead of replacing them.
- If user closes the terminal during runtime, there could be an invisible instance of Microsoft Word in the task manager that will need to be closed. Otherwise, one of the word documents will be impossible to delete because it is "in use".

## Requirements

- Windows Operating System
- Microsoft Word

- Windows Operating System
- Microsoft Word
- Python package manager (pip, conda, etc.)

## Installation

Nil.

## Updating Source Code

### Requirements

- Conda

### Instructions

1. Update source code found in `src`
2. Run `setup.bat` in either powershell or command prompt, note that the filepath to conda activate will need to be adjusted for this to work.. Terminal must have conda installed, check by running `conda --version`.
3. After installation, the `dist` folder will contain the tool.

## Future Improvements

- [ ] Add support for live editing within the same file, so that the user can see the changes being made as they are being made

##### Fix known issues

- [ ] Fix issue with not replacing words found in headers or footers
- [ ] Improve performance
  - [ ] Add support for multithreading
  - [ ] Add support for multiprocessing
- [ ] Fix issue with not replacing words found in shapes or text boxes
- [ ] Hide the terminal, or otherwise warn or prevent the user from stopping the process prematurely.
- [ ] Make the setup.bat work for all users. Currently, it only works for the Author's Miniconda installation pathfile.

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
- [x] Improve installation verbosity
  - [x] Add a progress bar
  - [x] Display more information in installation terminal during installation (via a INFO logger)
- [ ] Add a summary report of changes made (in the log or in a separate file)
