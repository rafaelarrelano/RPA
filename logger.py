"""
logger.py - Setup logging ke file dan console sekaligus
"""
import os
import logging
from datetime import datetime


def setup_logger() -> logging.Logger:
    log_dir = r"C:\RPA_StockRecon\Logs"
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(log_dir, f"rpa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

    logger = logging.getLogger("RPA_StockRecon")
    logger.setLevel(logging.DEBUG)

    # Format
    fmt = logging.Formatter("%(asctime)s | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S")

    # Handler: file
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    # Handler: console
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger
