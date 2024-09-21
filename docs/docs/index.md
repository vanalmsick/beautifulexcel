# BeautifulExcel

> **⚠️🏗️️ Note:**
> This is only the first version - actively working on additional features!

BeautifulExcel is a python package that makes it easy and quick to save pandas dataframes in beautifully formatted excel files. BeautifulExcel is the Openpyxl for Data Scientists with a deadline.
  
<br>

## Pandas vs. BeautifulExcel .to_excel()

| Pandas                                                                                                                                              | BeautifulExcel                                                                                                                                              |
|-----------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `df.to_excel(writer, sheet_name='My Output')`                                                                                                       | `writer.to_excel(df, sheet_name='My Output')`                                                                                                               |
| <img src="https://github.com/vanalmsick/beautifulexcel/raw/main/docs/docs/imgs/example_pandas.png" alt="Article Reading View" style="width:100%;"/> | <img src="https://github.com/vanalmsick/beautifulexcel/raw/main/docs/docs/imgs/example_beautifulexcel.png" alt="Article Reading View" style="width:100%;"/> |
| *<ins>Raw, unformatted</ins> data that requires lots of additional formatting.*                                                                     | *Quickly export <ins>beautifully</ins> styled table with only <ins>one line of code</ins>!*                                                                 |
  
<br>
  
## Getting it

```console
$ pip install beautifulexcel
```
**Update Package:** *(execute <ins>regularly</ins> to get the latest features)*
```console
$ pip install beautifulexcel --upgrade
```
  
<br>
  
## How to use:

```python
from beautifulexcel import ExcelWriter

with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as writer:
    ws1 = writer.to_excel(
        df,
        sheet_name='My Sheet',
        startrow=0,
        startcol=0,
        index=True,
        header=True,
        col_autofit=True,  # automatically change column width to fit content best
        col_widths={'A': 20, 'RoE': 40},  # define column width manually
        auto_number_formatting=True,  # automatically detect number format and change excel format
        style={'RoE': 'bg_light_blue', 'D:E': {'fill': 'FFEEB7'}},  # apply custom styling to this dataframe export
        use_theme_style=True,  # apply the excel workbook "theme" set in ExcelWriter()
    )
```
  
<br>

## Find out more about:
<div class="grid cards" markdown>

-   :material-file-excel-outline:{ .lg .middle } __beautifulexcel.ExcelWriter('workbook.xlsx')__

    ---

    Find out more about aguments for beautifulexcel.ExcelWriter(...)

    [:octicons-arrow-right-24: beautifulexcel.ExcelWriter()](./ExcelWriter.html#beautifulexcel.ExcelWriter.__init__)

-   :material-table:{ .lg .middle } __writer.to_excel(df, sheet_name='My Sheet')__

    ---

    Find out more about aguments for writer.to_excel(...)

    [:octicons-arrow-right-24: writer.to_excel()](./ExcelWriter.html#beautifulexcel.ExcelWriter.to_excel)

-   :material-format-font:{ .lg .middle } __Styling & Themes__

    ---

    Change the colors, fonts, sizes, borders and more with a few lines

    [:octicons-arrow-right-24: Customization](./styling.html)

-   :material-code-braces:{ .lg .middle } __Additional functions for sheets__

    ---

    Learn more about further functions to e.g. merge cells, group columns/rows and more

    [:octicons-arrow-right-24: sheet1.function_a()](./Sheet.html)

</div>

<br><br>