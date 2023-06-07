import csv
import logging
import click
from glossary_formatter import format_glossary


@click.command()
@click.argument("input_file", type=click.File("r"))
@click.argument("output_file", type=click.File("w"))
def main(input_docx, output_docx):
    """Formats a glossary CSV file into a human-readable format.

    Args:
        input_file (file): The input CSV file containing the glossary.
        output_file (file): The output file to write the formatted glossary to.

    Returns:
        None
    """
    # Set up logging
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler("app.log"), logging.StreamHandler()],
    )

    # Read input file and build glossary
    reader = csv.DictReader(
        input_docx,
        delimiter=",",
        fieldnames=[
            "file_path",
            "phrase_to_find",
            "font_style",
            "make_bold",
            "make_italic",
            "case_sensitive",
            "all_caps",
            "capitalize_first_letter",
            "highlight_color",
        ],
    )
    glossary = {}
    for row in reader:
        try:
            file_path = row["file_path"]
            phrase_to_find = row["phrase_to_find"]
            font_style = row["font_style"]
            make_bold = row["make_bold"]
            make_italic = row["make_italic"]
            case_sensitive = row["case_sensitive"]
            all_caps = row["all_caps"]
            capitalize_first_letter = row["capitalize_first_letter"]
            highlight_color = row["highlight_color"]
            glossary[file_path] = {
                "phrase_to_find": phrase_to_find,
                "font_style": font_style,
                "make_bold": make_bold,
                "make_italic": make_italic,
                "case_sensitive": case_sensitive,
                "all_caps": all_caps,
                "capitalize_first_letter": capitalize_first_letter,
                "highlight_color": highlight_color,
            }
        except KeyError:
            raise ValueError("Malformed input file")

    # Format glossary and write to output file
    formatted_glossary = format_glossary(glossary)
    output_docx.write(formatted_glossary)

    # Log success message
    logging.info("Glossary successfully formatted and written to output file")


if __name__ == "__main__":
    main()
