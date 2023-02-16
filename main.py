"""
Author: Hassan Mohamed
Documentation: Used to convert from excel files into CSV files, and extracts the quantity of the product.

# Discuss with Alex:
(OUTPUT)
-- Rounding in unit price.
-- Output columns Ordering.
"""

import logging
import pandas as pd
import sys
import re
import pathlib


def getQty(desc: str) -> str | None:
    """
    Extracts The Quantity from the Description of the product.

    Params:
        desc (str): The description of the product.

    Returns:
        Qty (str): The Quantiy of the product.
    """
    HasNumber = re.compile("\d")
    HasCorCS = re.compile("(\d*)C|CS$")

    quantity = None
    words = desc.split(" ")
    for word in words:
        if HasNumber.match(word) is None:
            continue
        if (res := HasCorCS.search(word)) is None:
            continue
        quantity = res.group(1)

    return quantity


def convert_sheet(df_input: pd.DataFrame, transaction: str):
    """
    The Main function tthat runs the program.
    """
    # Delete all rows before the data.
    header_num = 0
    for ind, row in df_input.iterrows():
        if list(row) == ["Description", "UPC #", "Quantity", "UOM", "Price", "Amount"]:
            header_num = ind
            break
    df_input = df_input.truncate(before=header_num)

    # Assign the correct columns' names and correct the index.
    df_input.rename(columns=df_input.iloc[0], inplace=True)
    df_input.drop(df_input.index[0], inplace=True)
    df_input.reset_index(inplace=True)

    # Create the output dataframe.
    output = pd.DataFrame(columns=["UPC #", "QTY", "Total Price", "Unit Cost"])
    # Populate the output dataframe.
    for row in df_input.itertuples():
        if (qty := getQty(row.Description)) is None:
            continue
        if pd.isna(row._3):
            logging.error(f"AT ROW: {row.Index}, UPC is Empty.")
        if pd.isna(row.Price):
            logging.error(
                f"AT ROW: {row.Index}, Price is not available. Will not include this entry in the output.")
            continue

        try:
            output.loc[len(output)] = [row._3,
                                       qty,
                                       row.Price,
                                       round(int(row.Price)/int(qty), 2),
                                       ]
        except Exception as e:
            logging.error(f"ERROR AT ROW: {row.Index}, {e}")

    # Save the output dataframe to disk.
    output.to_csv(f"./Output_{transaction}.csv", index=False)


if __name__ == "__main__":
    transactions = []
    for result in pathlib.Path(".").iterdir():
        if "transaction" in (result.name).lower() and "output" not in result.name:
            transactions.append(result)
    print(transactions)
    # open the work book and select the worksheet we want to convert.
    for transaction in transactions:
        workbook = pd.ExcelFile(transaction, "openpyxl")

        try:
            sheet = workbook.parse("Input")
        except Exception:
            logging.error("Please name the Input sheet 'Input' ")
            sys.exit()
        try:
            convert_sheet(sheet, transaction)
        except Exception as e:
            logging.error(e)
        finally:
            workbook.close()
