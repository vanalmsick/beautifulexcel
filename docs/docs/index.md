# BeautifulExcel

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
  
## Pandas vs. BeautifulExcel .to_excel()

| Pandas                                                                                                                                              | BeautifulExcel                                                                                                                                              |
|-----------------------------------------------------------------------------------------------------------------------------------------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `df.to_excel(writer, sheet_name='My Output')`                                                                                                       | `writer.to_excel(df, sheet_name='My Output')`                                                                                                               |
| <img src="https://github.com/vanalmsick/beautifulexcel/raw/main/docs/docs/imgs/example_pandas.png" alt="Article Reading View" style="width:100%;"/> | <img src="https://github.com/vanalmsick/beautifulexcel/raw/main/docs/docs/imgs/example_beautifulexcel.png" alt="Article Reading View" style="width:100%;"/> |
| *<ins>Raw, unformatted</ins> data that requires lots of additional formatting.*                                                                     | *Quickly export <ins>beautifully</ins> styled table with only <ins>one line of code</ins>!*                                                                 |
  
<br>
  
## How to use:

```python
from beautifulexcel import ExcelWriter

with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as writer:
    ws1 = writer.to_excel(df, sheet_name='My Sheet', startrow=0, startcol=0, index=True,
                          style={'RoE': 'bg_light_blue', 'D:E': {'fill': 'FFEEB7'}})
```
<br><br>