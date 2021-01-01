from dataclasses import dataclass
import re


@dataclass
class Item:
    """
    Finished good blueprint mapped from the source spreadsheet.
    """

    name: str
    level: str
    raw_material: str
    quantity: float
    unit: str

    def __post_init__(self):
        """
        Format the level property value to remove any extra dots.
        :return:
        """
        self.level = re.sub("[.]", "", str(self.level))
