import logging
import os

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

def setup_logger():
    logger = logging.getLogger()
    logger.setLevel(LOG_LEVEL)

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # File handler 
    file_handler = logging.FileHandler("app.log")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
