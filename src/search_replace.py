from pathlib import Path  # core python module
import win32com.client  # pip install pywin32
import csv
from typing import List, Tuple
import yaml


def read_config() -> dict:
    """
    Reads configuration data from a YAML file.

    Returns:
        dict: A dictionary containing the configuration data.
    """
    current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
    config_file = current_dir / "config.yml"

    with open(config_file, "r") as f:
        config_data = yaml.safe_load(f)

    return config_data


def read_csv_file(
    csv_directory_name: str = "search_replace_pairs",
) -> List[Tuple[str, str]]:
    """
    Reads a CSV file containing search and replace pairs.

    Args:
        filepath (str): The path to the CSV file.

    Returns:
        List[Tuple[str, str]]: A list of tuples containing the search and replace phrases.
    """
    current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
    csv_dir = current_dir / csv_directory_name

    search_replace_pairs = []
    for csv_file in Path(csv_dir).rglob("*.csv*"):
        with open(csv_file, "r") as csvfile:
            reader = csv.reader(csvfile)
            next(reader)  # skip the first row
            for row in reader:
                search_replace_pairs.append((row[0], row[1]))
    return search_replace_pairs


def replace_words_in_word_document(
    search_replace_pairs: List[Tuple[str, str]],
    replace_all_or_first_word: int = 2,
    input_dir_name: str = "input",
    output_dir_name: str = "output",
) -> None:
    """
    Replaces all occurrences of multiple strings in all Word documents in a directory.

    Args:
        search_replace_pairs (List[Tuple[str, str]]): A list of tuples containing the search and replace phrases.
        replace_all_or_first_word (int): 2=replace all occurances, 1=replace one occurence, 0=replace no occurences. Default is 2.
    """

    # current path
    current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()

    # filepaths of input and output directories containing word documents
    input_dir = current_dir / input_dir_name
    output_dir = current_dir / output_dir_name

    # create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    # Open Word
    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    for doc_file in Path(input_dir).rglob("*.doc*"):
        # Open each document and replace strings
        word_app.Documents.Open(str(doc_file))

        for search_str, replace_str in search_replace_pairs:
            # API documentation: https://learn.microsoft.com/en-us/office/vba/api/word.find.execute
            word_app.Selection.Find.Execute(
                FindText=search_str,
                ReplaceWith=replace_str,
                Replace=replace_all_or_first_word,
                Forward=True,
                MatchCase=True,
                MatchWholeWord=False,
                MatchWildcards=True,
                MatchSoundsLike=False,
                MatchAllWordForms=False,
                Wrap=0,
                Format=True,
            )

            # -- Replace str in shapes
            # VBA SO reference: https://stackoverflow.com/a/26266598
            # Loop through all the shapes
            for i in range(word_app.ActiveDocument.Shapes.Count):
                if word_app.ActiveDocument.Shapes(i + 1).TextFrame.HasText:
                    words = word_app.ActiveDocument.Shapes(
                        i + 1
                    ).TextFrame.TextRange.Words
                    # Loop through each word. This method preserves formatting.
                    for j in range(words.Count):
                        # If a word exists, replace the text of it, but keep the formatting.
                        if (
                            word_app.ActiveDocument.Shapes(i + 1)
                            .TextFrame.TextRange.Words.Item(j + 1)
                            .Text
                            == search_str
                        ):
                            word_app.ActiveDocument.Shapes(
                                i + 1
                            ).TextFrame.TextRange.Words.Item(j + 1).Text = replace_str

        # Save the new file
        output_path = output_dir / f"{doc_file.stem}_replaced{doc_file.suffix}"
        word_app.ActiveDocument.SaveAs(str(output_path))
        word_app.ActiveDocument.Close(SaveChanges=False)
    word_app.Application.Quit()


if __name__ == "__main__":
    config_data = read_config()
    search_replace_pairs = read_csv_file(
        csv_directory_name=config_data.get("CSV_DIRECTORY_NAME", "search_replace_pairs")
    )
    replace_words_in_word_document(
        search_replace_pairs,
        replace_all_or_first_word=config_data.get("REPLACE_ALL_OR_FIRST_WORD", 2),
        input_dir_name=config_data.get("INPUT_DIRECTORY_NAME", "input"),
        output_dir_name=config_data.get("OUTPUT_DIRECTORY_NAME", "output"),
    )
