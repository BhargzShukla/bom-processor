#!/usr/bin/env python

import pandas as pd

from src.utils import (
    get_data_from_source_sheet,
    load_data_from_file,
    get_item_groups,
    flatten,
    custom_sheet,
)

DATA_FILE = "data/data_clean.xlsx"

source_sheet, source_workbook = load_data_from_file(DATA_FILE)
source_data = get_data_from_source_sheet(source_sheet)
processed_items = flatten(source_data)

data = pd.DataFrame(processed_items)
item_groups = get_item_groups(data)


def write_processed_data_to_sheet(processed_data):
    for idx, (key, val) in enumerate(processed_data.items(), start=1):
        sheet = custom_sheet(workbook=source_workbook, title=key, key=key)
        for item_idx, item_val in enumerate(list(val), start=1):
            item_name = data.iloc[item_val]["raw_material"]
            item_quantity = data.iloc[item_val]["quantity"]
            item_unit = data.iloc[item_val]["unit"]
            sheet.append([item_idx, item_name, item_quantity, item_unit])
        sheet.append(["End of RM"])
    source_workbook.save(DATA_FILE)


write_processed_data_to_sheet(item_groups)
