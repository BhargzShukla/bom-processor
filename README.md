# BOM Parser

A Python script to parse a source for raw materials and create the necessary BOMs for finished goods.

Dependencies:

- `pandas`
- `openpyxl`

Dev dependencies:

- `ipython`
- `black`
- `pre-commit`

## Brief Intro

A BOM (Bill of Material) is a document detailing the amount of raw materials required to make something. Typically used
in manufacturing, BOMs can give a flattened view into what your manufacturing process looks like in terms of raw
materials required. For instance, here's a multi-level BOM for a pencil:

- ID: PENC1L
- Finished Good: Pencil - 1 Pc
- Raw Materials:
    1. Clay binder | 1 Pc
        1. Clay | 10 Gm
        2. Wood | 1 Pc
    2. Graphite | 1 Pc

## Task List

1. Parse provided excel sheet
2. Prepare individual BOM sheets for each item in the source spreadsheet

## Solution approach

Flatten the source spreadsheet into level 1 items, group the resulting dataset according to the item names, and create
individual spreadsheets for each item.

## Run the project

To run this project locally, follow these steps:

1. Clone the repo
2. Install dependencies from `requirements.txt`
3. Run `main.py`

There is some data provided in the `data` folder. Out of the 2 files, `data.xlsx` contains the source data as well as a
couple of examples of the required BOM sheets. `data_clean.xlsx` is a copy of `data.xlsx`, with the sample BOM sheets
removed.

To test with your own data, modify the `DATA_FILE` variable (line 7) in the `main.py` script.

## Possible Enhancements

1. Implement a folder based queue system which scans and/or watches an input folder for spreadsheets to process them to
   an output folder.
2. Parse source spreadsheet into a tree data structure.
