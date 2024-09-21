# Getting Started

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

    [:octicons-arrow-right-24: beautifulexcel.ExcelWriter()](./ExcelWriter/#beautifulexcel.ExcelWriter.__init__)

-   :material-table:{ .lg .middle } __writer.to_excel(df, sheet_name='My Sheet')__

    ---

    Find out more about aguments for writer.to_excel(...)

    [:octicons-arrow-right-24: writer.to_excel()](./ExcelWriter/#beautifulexcel.ExcelWriter.to_excel)

-   :material-format-font:{ .lg .middle } __Styling & Themes__

    ---

    Change the colors, fonts, sizes, borders and more with a few lines

    [:octicons-arrow-right-24: Customization](./styling/)

-   :material-code-braces:{ .lg .middle } __Additional functions for sheets__

    ---

    Learn more about further functions to e.g. merge cells, group columns/rows and more

    [:octicons-arrow-right-24: sheet1.function_a()](./Sheet/)

</div>

<br>