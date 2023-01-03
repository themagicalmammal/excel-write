# Excel Write
[![DeepSource](https://deepsource.io/gh/themagicalmammal/excel-write.svg/?label=active+issues&show_trend=true)](https://deepsource.io/gh/themagicalmammal/excel-write/?ref=repository-badge)
[![DeepSource](https://deepsource.io/gh/themagicalmammal/excel-write.svg/?label=resolved+issues&show_trend=true)](https://deepsource.io/gh/themagicalmammal/excel-write/?ref=repository-badge)


Optimised way to write in excel files.

Developed by [Dipan Nanda](https://github.com/themagicalmammal) (c) 2023

## Example of Usage

### write_in_excel

```python
from excel_write import write_in_excel

write_in_excel(df, location, sheet)
"""
:param DataFrame df: The DataFrame used to export to Excel
:param str sheet: The name that is to be assigned to the file
:param str location: Location where the file is to be created
:param bool index: including index or not
"""
```

### auto_adjust_excel_width

```python
from excel_write import auto_adjust_excel_width

auto_adjust_column_width_index(df, writer, sheet_name="MySheet", margin=3)

"""
:param DataFrame df: The DataFrame used to export the Excel
:param pd.ExcelWriter writer: The pandas exporter with engine="xlsxwriter"
:param str sheet_name: The name of the sheet
:param int margin: How many extra space (beyond the maximum size of the string)
:param int length_factor: The factor to apply to the character length to obtain the 
column width
:param int decimals: The number of decimal places to assume for floats: Should be the
same as the number of decimals displayed in the Excel
:param bool index: Whether the DataFrame's index is inserted as a separate column (if
index=False in df.to_xlsx() set index=False here!)
"""
```


## Changelog
Go [here](CHANGELOG.md) to checkout the complete changelog.

## License
#### This is under MIT License
[![License: MIT](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
