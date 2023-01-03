#!/usr/bin/env python3
"""Description: Simple way to write in Excel"""

import functools
import os.path
from decimal import Decimal

import openpyxl.utils.cell
from pandas import ExcelWriter


def write_in_excel(df, location, sheet, index=False):
    """
    Writes pandas DataFrame in Excel depending on the state of the file

    How to use:
    ```
    write_in_excel(df, location, sheet)
    ```
    :param DataFrame df: The DataFrame used to export to Excel
    :param str sheet: The name that is to be assigned to the file
    :param str location: Location where the file is to be created
    :param bool index: including index or not
    """
    try:
        if not os.path.isfile(location):
            # pylint: disable=abstract-class-instantiated
            writer = ExcelWriter(location, engine="xlsxwriter")
            df.to_excel(writer, sheet_name=sheet, index=index)
            auto_adjust_excel_width(df, writer, sheet_name=sheet, margin=0)
            worksheet = writer.sheets[sheet]  # pull worksheet object
            worksheet.freeze_panes(1, 0)
            for idx, col in enumerate(df):  # loop through all columns
                series = df[col]
                max_len = (
                    max((
                        series.astype(str).map(
                            len).max(),  # len of largest item
                        len(str(series.name)),  # len of column name/header
                    )) + 1)  # adding a little extra space
                max_len = min(max_len, 50)
                worksheet.set_column(idx, idx, max_len)  # set column width
            writer.close()
        else:
            # pylint: disable=abstract-class-instantiated
            with ExcelWriter(location,
                             mode="a",
                             engine="openpyxl",
                             if_sheet_exists="replace") as writer:
                df.to_excel(writer, sheet_name=sheet, index=index)
                auto_adjust_excel_width(df, writer, sheet_name=sheet, margin=0)

    except PermissionError:
        # Helps in case if the Excel is already in access mode somewhere else
        print(f"Failed to save {location} : Try closing excel doc")
    except Exception as e:
        print(e)


def find_length(text):
    """
    Get the effective text length in characters, taking into account newlines

    How to use:
    ```
    find_length(text)
    ```
    :param str text: The text we are checking the length for
    """
    if not text:
        return 0
    lines = text.split("\n")
    return max(len(line) for line in lines)


def find_float_length(v, decimals=3):
    """
    Like str() but rounds decimals to predefined length

    How to use:
    ```
    find_float_length(v, decimals)
    ```
    :param float v: The float value we will check the length for
    :param int decimals: The amount of decimal points
    """
    if isinstance(v, float):  # Round to [decimal] places
        return str(
            Decimal(v).quantize(Decimal("1." + "0" * decimals)).normalize())
    return str(v)


def auto_adjust_excel_width(df,
                            writer,
                            sheet_name,
                            margin=3,
                            length_factor=1.0,
                            decimals=3,
                            index=True):
    """
    Auto adjust column width to fit content in a XLSX exported from a pandas DataFrame.

    How to use:
    ```
    with ExcelWriter(filename) as writer:
    df.to_excel(writer, sheet_name="MySheet")
    auto_adjust_column_width_index(df, writer, sheet_name="MySheet", margin=3)
    ```

    :param DataFrame df: The DataFrame used to export the Excel
    :param ExcelWriter writer: The pandas exporter with engine="xlsxwriter"
    :param str sheet_name: The name of the sheet
    :param int margin: How many extra space (beyond the maximum size of the string)
    :param int length_factor: The factor to apply to the character length to obtain
    the column width
    :param int decimals: The number of decimal places to assume for floats: Should be
    the same as the number of decimals displayed in the Excel
    :param bool index: Whether the DataFrame's index is inserted as a separate column
    (if index=False in df.to_xlsx()
    set index=False here!)
    """
    writer_type = type(
        writer.book
    ).__module__  # e.g. 'xlsxwriter.workbook' or 'openpyxl.workbook.workbook'
    is_openpyxl = writer_type.startswith("openpyxl")
    is_xlsxwriter = writer_type.startswith("xlsxwriter")
    to_str = functools.partial(find_float_length, decimals=decimals)
    # str() but rounds decimals to predefined length
    if not is_openpyxl and not is_xlsxwriter:
        raise ValueError(
            "Only openpyxl and xlsxwriter are supported as backends, not " +
            writer_type)
    sheet = writer.sheets[sheet_name]
    # Compute & set column width for each column
    for column_name in df.columns:
        # Convert the value of the columns to string and select the
        column_length = max(
            df[column_name].apply(to_str).map(find_length).max(),
            find_length(column_name),
        )
        # Get index of column in Excel
        # Column index is +1 if we also export the index column
        col_idx = df.columns.get_loc(column_name)
        if index:
            col_idx += 1
        # Set width of column to (column_length + margin)
        if is_openpyxl:
            sheet.column_dimensions[openpyxl.utils.cell.get_column_letter(
                col_idx + 1)].width = (column_length * length_factor + margin)
        else:
            sheet.set_column(col_idx, col_idx,
                             column_length * length_factor + margin)
    if index:  # If the index column is being exported
        index_length = max(
            df.index.map(to_str).map(find_length).max(),
            find_length(df.index.name))
        if is_openpyxl:
            sheet.column_dimensions[
                "A"].width = index_length * length_factor + margin
        else:
            sheet.set_column(0, 0, index_length * length_factor + margin)
