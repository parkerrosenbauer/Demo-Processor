from __future__ import annotations
import json
import datetime
from logs.log import logger

VALIDATION_COUNTS = r"\\CT-FS10\BLR_Share\Marketing\_Database " \
                    r"Management\PRosenbauer\demo_automation\data\validation_counts.json "


class Validation:
    def __init__(self, demo_type, demo_date):
        """Initialize Validation.

        :param demo_date: date of the demo
        :type demo_date: datetime.datetime
        """
        self.idx = f"{demo_type} ({demo_date.strftime('%#m/%#d/%Y')})"

    def update_counts(self, **kwargs) -> None:
        """Update the demo specific variables.

        :param kwargs: variable names
        :type kwargs: any
        :return: None
        :rtype: None
        """
        with open(VALIDATION_COUNTS, 'r+') as file:
            counts = json.load(file)

            for item, value in kwargs.items():
                if item in counts[self.idx]:
                    counts[self.idx][item] = value
                    logger.info("count item: %s, count: %s", item, str(value))
                else:
                    raise ValueError(f"{item} is not a valid metric.")

            file.seek(0)
            json.dump(counts, file)
            file.truncate()

    def retrieve_one(self, item: str) -> int | dict | list:
        """Retrieve a demo specific variable.

        :param item: variable name
        :type item: str
        :return: variable value
        :rtype: int | dict | list
        """
        with open(VALIDATION_COUNTS, 'r') as file:
            counts = json.loads(file.read())
        return counts[self.idx][item]

    def retrieve_all(self) -> dict:
        """Retrieve all demo specific variables.

        :return: dictionary of all variables
        :rtype: dict
        """
        with open(VALIDATION_COUNTS, 'r') as file:
            counts = json.loads(file.read())
        return counts[self.idx]
