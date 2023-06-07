import logging
from logging import FileHandler


LOG_LEVEL: str = "INFO"
SAVE_LOG_TO_FILE: bool = True
LOG_FILE_NAME: str = "log.log"
LOG_FORMAT: str = "%(asctime)s:%(levelname)s:%(message)s"


def create_logger(
    log_level: str = LOG_LEVEL,
    log_file_name: str = LOG_FILE_NAME,
    save_log_to_file: bool = SAVE_LOG_TO_FILE,
    log_format: str = LOG_FORMAT,
) -> logging.Logger:
    """
    Creates a logger with a console handler and an optional file handler.

    Args:
        log_level (str): The logging level for the logger. Default is set to "INFO".
        log_file_name (str): The name of the log file to be created. Default is set to "app.log".
        save_log_to_file (bool): Whether or not to create a file handler for the logger. Default is set to True.
        log_format (str): The format of the log messages. Default is set by LOG_FORMAT
    Returns:
        A logging.Logger object.
    """
    logger = logging.getLogger()
    log_format = logging.Formatter(LOG_FORMAT)
    logger.setLevel(log_level)

    # create a console handler
    console_handler: logging.StreamHandler = logging.StreamHandler()
    console_handler.setFormatter(log_format)
    logger.addHandler(console_handler)

    if save_log_to_file:
        try:
            # create a file handler
            file_handler = FileHandler(log_file_name)
            file_handler.setFormatter(log_format)
            logger.addHandler(file_handler)
        except PermissionError as error:
            print(
                f"Permission denied to create log file {log_file_name}. Check permissions for editing the directory."
            )
            raise error
        except Exception as error:
            print(f"Error creating file handler: {error}")
            raise error

    return logger
