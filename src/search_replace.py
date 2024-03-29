# %%
# IMPORTS

from pathlib import Path  # core python module
import win32com.client  # pip install pywin32
import csv
from typing import List, Tuple
import yaml
from tqdm import tqdm
import logging

# import kivy


# %%
# LOGGER


def _setup_logger(
    logger_name: str = "log",
    log_file: str = "log.log",
    level: str = "INFO",
    file_handler_on: bool = False,
    simple_format: bool = False,
) -> logging.Logger:
    # Create a logger object
    logger = logging.getLogger(logger_name)

    # Set the logging level
    logger.setLevel(logging.getLevelName(level.upper()))

    # Create a formatter that formats log messages
    if simple_format is True:
        formatter = logging.Formatter("%(message)s")
    else:
        formatter = logging.Formatter(
            "%(asctime)s -  %(levelname)s - %(funcName)s - (Line: %(lineno)d) - %(message)s"
        )

    # Create a console handler that logs messages to the console
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # Create a file handler that logs messages to a file if file_handler_on is True
    if file_handler_on:
        file_handler = logging.FileHandler(log_file)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger


logger = _setup_logger(simple_format=True, file_handler_on=True)
logger.info("Let's rock & roll...")
logger = _setup_logger(level="DEBUG", file_handler_on=True, simple_format=False)

# %%
#


def _read_config() -> dict:
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


def _read_csv_file(
    csv_directory_name: str,
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


# %%
# COMPOSABLE FUNCTIONS


def _replace_words_in_word_document(
    search_replace_pairs: List[Tuple[str, str]],
    replace_all_or_first_word: int,
    input_dir_name: str,
    output_dir_name: str,
) -> None:
    """
    Replaces all occurrences of multiple strings in all Word documents in a directory.

    Args:
        search_replace_pairs (List[Tuple[str, str]]): A list of tuples containing the search and replace phrases.
        replace_all_or_first_word (int): 2=replace all occurrances, 1=replace one occurrence, 0=replace no occurrences. Default is 2.
    """

    # current path
    current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()

    # filepaths of input and output directories containing word documents
    input_dir = current_dir / input_dir_name
    output_dir = current_dir / output_dir_name

    # create output directory if it doesn't exist
    output_dir.mkdir(parents=True, exist_ok=True)

    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    for doc_file in tqdm(list(Path(input_dir).rglob("*.doc*"))):
        # Open Word
        logger.debug("Opening Word.")

        try:
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

                # -- Replace str in headers
                logger.debug(
                    f"For replacement pair ({len(search_replace_pairs)}) in '{doc_file.name}, replacing '{search_str}' with '{replace_str}' )."
                )

            # -- Replace str in shapes
            # VBA SO reference: https://stackoverflow.com/a/26266598
            # Loop through all the shapes
            try:
                number_of_shapes_in_document = word_app.ActiveDocument.Shapes.Count

                for i in range(number_of_shapes_in_document):
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
                                logger.debug(
                                    f"For shape ({i+1}/{number_of_shapes_in_document}), replacing '{search_str}' with '{replace_str}'."
                                )
                                word_app.ActiveDocument.Shapes(
                                    i + 1
                                ).TextFrame.TextRange.Words.Item(
                                    j + 1
                                ).Text = replace_str
                                logger.debug("Replaced.")
            except AttributeError:
                logger.debug(f"No shapes in document: {doc_file.name}.")
            except Exception as e:
                logger.warning(
                    f"Problem processing text in shapes for {doc_file.name}. Occurred at shape ({i+1}/{number_of_shapes_in_document}) when replacing '{search_str}' with '{replace_str}': {e}"
                )

        except Exception as e:
            logger.warning(f"Problem processing document {doc_file.name}: {e}")

        # Save the new file
        logger.debug(f"Saving the new file for {doc_file.name}")
        output_path = output_dir / f"{doc_file.stem}_replaced{doc_file.suffix}"
        word_app.ActiveDocument.SaveAs(str(output_path))
        logger.debug(f"Saved the new file for {doc_file.name}.")

        # Close the document
        logger.debug(f"Closing the document {doc_file.name}.")
        word_app.ActiveDocument.Close(SaveChanges=False)
        logger.debug(f"Closed the document {doc_file.name}.")

    # Quit Word
    logger.debug("Quitting Word.")
    word_app.Application.Quit()


# %%
# MODULE LEVEL FUNCTION


def search_replace() -> None:
    # Read config file
    logger.info("Searching config.yml for configuration data.")
    config_data = _read_config()

    # Read CSV file containing search and replace pairs
    logger.info("Reading CSV file containing search and replace pairs.")
    search_replace_pairs = _read_csv_file(
        csv_directory_name=config_data.get(
            "CSV_DIRECTORY_NAME", "2_insert_search-replace_pairs_here"
        )
    )

    # Replace words in Word documents
    logger.info("Replacing words in Word documents.")
    _replace_words_in_word_document(
        search_replace_pairs,
        replace_all_or_first_word=config_data.get("REPLACE_ALL_OR_FIRST_WORD", 2),
        input_dir_name=config_data.get(
            "INPUT_DIRECTORY_NAME", "1_insert_word_documents_here"
        ),
        output_dir_name=config_data.get("OUTPUT_DIRECTORY_NAME", "3_output"),
    )

    # Done! message
    logger.info("Done!")


# %%
# __name__ == "main":
if __name__ == "__main__":
    search_replace()
