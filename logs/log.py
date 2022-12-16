import os
import sys
import logging

log_format = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s", datefmt='%m/%d/%Y')
logger = logging.getLogger()

log_file = os.path.join(os.path.dirname(__file__), 'log.txt')

file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(log_format)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(log_format)
logger.addHandler(console_handler)

logger.setLevel(logging.INFO)
