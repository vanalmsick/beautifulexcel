# BeautifulExcel

<img src="https://github.com/vanalmsick/beautifulexcel/raw/main/docs/docs/imgs/social.png" alt="Social Banner" style="width:100%;"/>

> **⚠️🏗️️ Note:**
> This is only the first version - actively working on additional features!

BeautifulExcel is a python package that makes it easy and quick to save pandas dataframes in beautifully formatted excel files. BeautifulExcel is the Openpyxl for Data Scientists with a deadline.

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
  
## Further details in the [Documentation](https://vanalmsick.github.io/beautifulexcel/)
[Go to **Documentation**](https://vanalmsick.github.io/beautifulexcel/)
  
<br>
  
## A feature is missing? Feel free to contribute!
- Please submit new features as Pull Request to the "dev" branch
- Please make sure the code is nicely formatted and has doc strings by executing `$ pre-commit install` before your git commit
