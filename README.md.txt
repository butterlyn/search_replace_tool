

# Search and Replace

This Python script replaces all occurrences of multiple strings in all Word documents in a directory. It reads one or multiple CSV files containing search and replace pairs, and replaces all occurrences of the search phrases with the corresponding replace phrases in all Word documents in the input directory. The modified documents are saved in the output directory.

## Requirements

- Python 3.x
- Microsoft Word

## Installation



## Usage

1. Place the Word documents to be modified in the `input` directory.
2. Create a CSV file containing the search and replace pairs. The first row should contain the column headers "Search" and "Replace". Subsequent rows should contain the search and replace phrases, respectively.
3. Run the script with `python search_replace.py`.
4. The modified Word documents will be saved in the `output` directory.

## Configuration

The following configuration options are available:

- `wd_replace`: 2=replace all occurrences, 1=replace one occurrence, 0=replace no occurrences. Default is 2.
- `wd_find_wrap`: 2=ask to continue, 1=continue search, 0=end if search range is reached. Default is 0.