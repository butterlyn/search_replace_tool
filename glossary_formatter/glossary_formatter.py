from typing import Optional
from typing import List
from glossary_formatter import create_logger
import warnings
import win32com

# initialise logger
logger = create_logger.create_logger()

try:
    import win32com.client.constants as constants
    import win32com.client
except ImportError:
    warnings.warn(
        "win32com.client module not found. Install pywin32 package to use this module."
    )


def open_docx_file(file_path: str) -> win32com.client.CDispatch:
    """Open a Word document and return the document object.

    Args:
        file_path (str): The path to the Word document.

    Returns:
        win32com.client.CDispatch: The Word document object.

    """
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(file_path)
    return doc


def find_occurrences(
    doc: win32com.client.CDispatch, phrase_to_find: str, case_sensitive: bool
) -> List[str]:
    """Find all occurrences of a phrase in a Word document.

    Args:
        doc (win32com.client.CDispatch): The Word document object.
        phrase_to_find (str): The phrase to search for in the document.
        case_sensitive (bool): Whether or not the search should be case-sensitive.

    Returns:
        List[str]: A list of found occurrences.
    """
    found_occurrences = []
    search_range = doc.Range()
    search_range.Find.ClearFormatting()
    search_range.Find.MatchCase = case_sensitive
    search_range.Find.Text = phrase_to_find
    while search_range.Find.Execute():
        found_occurrences.append(search_range.Text)
    return found_occurrences


def apply_font_style(
    search_range: win32com.client.CDispatch, font_style: str
) -> List[str]:
    """Apply the specified font style to the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.
        font_style (str): The font style to apply to the found text.

    Returns:
        List[str]: A list of changes made.

    """
    changes_made = []
    try:
        search_range.Font.Name = font_style
        changes_made.append(f"Font style set to '{font_style}'")
    except AttributeError:
        logger.error(
            f"Invalid font style: {font_style}. Possible font styles are: {win32com.client.constants.FontNames}"
        )
    return changes_made


def apply_bold_style(search_range: win32com.client.CDispatch) -> List[str]:
    """Apply bold font style to the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.

    Returns:
        List[str]: A list of changes made.

    """
    search_range.Font.Bold = True
    return ["Bold font style applied"]


def apply_italic_style(search_range: win32com.client.CDispatch) -> List[str]:
    """Apply italic font style to the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.

    Returns:
        List[str]: A list of changes made.

    """
    search_range.Font.Italic = True
    return ["Italic font style applied"]


def apply_all_caps(search_range: win32com.client.CDispatch) -> List[str]:
    """Apply all caps to the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.

    Returns:
        List[str]: A list of changes made.

    """
    search_range.Font.AllCaps = True
    return ["All caps applied"]


def capitalize_first_letter(search_range: win32com.client.CDispatch) -> List[str]:
    """Capitalize the first letter of the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.

    Returns:
        List[str]: A list of changes made.

    """
    search_range.Case = constants.wdTitleWord
    return ["First letter capitalized"]


def apply_highlight_color(
    search_range: win32com.client.CDispatch, highlight_color: str
) -> List[str]:
    """Apply the specified highlight color to the found text.

    Args:
        search_range (win32com.client.CDispatch): The range object representing the found text.
        highlight_color (str): The color to use for highlighting the found text.

    Returns:
        List[str]: A list of changes made.

    """
    valid_colors = [
        "none",
        "yellow",
        "green",
        "cyan",
        "magenta",
        "blue",
        "red",
        "darkblue",
        "darkred",
        "darkyellow",
        "darkgreen",
        "darkcyan",
        "darkmagenta",
        "darkgray",
    ]
    changes_made = []
    if highlight_color.lower() in valid_colors:
        search_range.HighlightColorIndex = getattr(
            constants, f"wd{highlight_color.capitalize()}"
        )
        changes_made.append(f"Highlight color set to '{highlight_color}'")
    else:
        logger.error(
            f"Invalid highlight color: {highlight_color}. Valid highlight colors are: {', '.join(valid_colors)}. Skipping highlight color."
        )
    return changes_made


def save_docx_file(
    doc: win32com.client.CDispatch,
    found_occurrences: List[str],
    changes_made: List[str],
    file_path: str,
    phrase_to_find: str,
) -> None:
    """Save the Word document and log the changes made.

    Args:
        doc (win32com.client.CDispatch): The Word document object.
        found_occurrences (List[str]): A list of found occurrences.
        changes_made (List[str]): A list of changes made.
        file_path (str): The path to the Word document.
        phrase_to_find (str): The phrase that was searched for in the document.

    """
    if found_occurrences:
        doc.Save()
        logger.info(
            f"Found and styled {len(found_occurrences)} occurrences of '{phrase_to_find}' in {file_path}."
        )
        if changes_made:
            logger.info(f"Changes made to {file_path}: {', '.join(changes_made)}")
    else:
        logger.warning(f"Phrase '{phrase_to_find}' not found in {file_path}.")


def format_glossary(
    file_path: str,
    phrase_to_find: str,
    font_style: Optional[str] = None,
    make_bold: bool = False,
    make_italic: bool = False,
    case_sensitive: bool = False,
    all_caps: bool = False,
    capitalize_first_letter: bool = False,
    highlight_color: Optional[str] = None,
) -> None:
    """Find and style all occurrences of a phrase in a Word document.

    Args:
        file_path (str): The path to the Word document.
        phrase_to_find (str): The phrase to search for in the document.
        font_style (Optional[str], optional): The font style to apply to the found text. Defaults to None.
        bold (bool, optional): Whether or not to apply bold font style to the found text. Defaults to False.
        italic (bool, optional): Whether or not to apply italic font style to the found text. Defaults to False.
        case_sensitive (bool, optional): Whether or not the search should be case-sensitive. Defaults to False.
        all_caps (bool, optional): Whether or not to apply all caps to the found text. Defaults to False.
        capitalize_first_letter (bool, optional): Whether or not to capitalize the first letter of the found text. Defaults to False.
        highlight_color (Optional[str], optional): The highlight color to apply to the found text. Defaults to None.
    """
    doc = open_docx_file(file_path)
    found_occurrences = find_occurrences(doc, phrase_to_find, case_sensitive)
    changes_made = []
    for occurrence in found_occurrences:
        search_range = doc.Range()
        search_range.Find.ClearFormatting()
        search_range.Find.MatchCase = case_sensitive
        search_range.Find.Text = occurrence
        if font_style:
            apply_font_style(search_range, font_style)
        if make_bold:
            apply_bold_style(search_range)
        if make_italic:
            apply_italic_style(search_range)
        if all_caps:
            apply_all_caps(search_range)
        if capitalize_first_letter:
            capitalize_first_letter(search_range)
        if highlight_color:
            apply_highlight_color(search_range, highlight_color)
        changes_made.append(occurrence)
    save_docx_file(doc, found_occurrences, changes_made, file_path, phrase_to_find)
