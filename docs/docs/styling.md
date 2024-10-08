# How to apply cell sytling?

### **Style: *Add "style"-ing to individual dataframe exports***

In `writer.to_excel(df, ..., style={})` you can define specific styling kwargs for that specific table.  
The **style dictionary** syntax is:

| dictionary key:<br>reference the column, row, or cell                                                                                                                                                                                                                                                                                                                                                       | dictionary value:<br>provide formatting specs                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Options:<br><ul><li>***df column name***<br><ul><li>single coumn 'emplyees' or</li><li>range 'inception:last_contact'</li></ul></li><li>***df row number***<br><ul><li>single row '1' or</li><li>range '1:5'</li></ul></li><li>***excel column***<ul><li>single cell 'A1' or</li><li>range 'A1:C3'</li></ul></li><li>***excel column***<ul><li>single column 'A' or</li><li>range 'A:C'</li></ul></li></ul> | Options:<br><ul><li>***preset name*** from the selected "theme" *([see presets of 'elegant_blue'](https://github.com/vanalmsick/beautifulexcel/blob/main/beautifulexcel/themes/elegant_blue.yml))*<ul><li>single preset 'bg_light_blue' or</li><li>list of presets ['bg_light_blue', 'num_fmt_pct']</li></li></ul></li><li>***custom stying kwargs*** as dictionary as per [*openpyxl's class names*](https://openpyxl.readthedocs.io/en/stable/styles.html); examples:<ul><li>_font\_\_name: 'Arial'_</li><li>_font\_\_size: 10_</li><li>_font\_\_bold: True_</li><li>fill: 'FFEEB7'</li><li>_alignment\_\_horizontal: 'center'_</li><li>_alignment\_\_vertical: 'center'_</li><li>_numberformat: '#,##0'_</li><li>...</li></ul></li></ul> |


**Examples:** _(showcasing the many different styling options)_

```python
style = {'emplyees': ['bg_light_blue', 'num_fmt_pct'], 'F:G': 'num_fmt_pct'}
```

```python
style = {'C3:D10': {'font__size': 20, 'numberformat': '#,##0', 'font__italic'=True}, 'employees:customers': {'numberformat': '#,##0'}}
```

```python
MY_CUSTOM_WARNING_STYLE = {'font__bold': True, 'text__color': 'ff0000', 'font__size': 20}
MY_CUSTOM_DATE_STYLE = {'numberformat': 'yyyy-mm-dd'}

style = {
  '1': MY_CUSTOM_WARNING_STYLE, 
  '2:5': {'font__size': 20},
  'B3:G10': ['bg_light_blue', 'num_fmt_pct'], 
  'A1': {**MY_CUSTOM_WARNING_STYLE, **MY_CUSTOM_DATE_STYLE}
}
```

### **Theme: *Set "theme" for entire excel file***

In `ExcelWriter(..., theme='elegant_blue')` you can define the base theme that will be applied to your entire Excel file.  
You can pass either:

- a ***theme name*** like 'elegant_blue',
- or your personal ***.yml-theme-file path*** _([syntax example here](https://github.com/vanalmsick/beautifulexcel/blob/main/beautifulexcel/themes/elegant_blue.yml))_
  
<br>