# beautifulexcel.ExcelWriter()

> **beautifulexcel.ExcelWriter**(***file**='excel.xlsx', **mode**='replace', **if_sheet_exists**='replace', **theme**='elegant_blue', **ref_warnings**=True, **date_format**=None, **datetime_format**=None, **engine_kwargs**={}*)

> ### Parameters:
> - **file: *str***<br>File path
> - **mode: *"replace", "modify", default "replace"***<br>Excel Workbook behavious if file already exists.
> - **if_sheet_exists: *"error", "new", "replace", "overlay", default "replace"***<br>Behaviour if excel sheet already exists
> - **theme: *str path or pre-defined theme name***<br>Provide path to personal yml-theme-file or use name of pre-defined theme "elegant_blue", "elegant_green", "elegant_yellow" 
> - **ref_warnings: *bool, default True***<br>By setting parameter to False you can supress cell ref warnings of not found cell refs
> - **date_format: *str, default None***
> - **datetime_format: *str, default None***
> - **engine_kwargs: *dict, default {}*** 

> ### Examples:
> ```python
> from beautifulexcel import ExcelWriter
> 
> with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as writer:
>     ...
> ```

> ### Methods:
> *Save pandas.DataFrame to excel:*  
> 
> **writer.to_excel**(  
> &nbsp;&nbsp;&nbsp;&nbsp;**df**,  
> &nbsp;&nbsp;&nbsp;&nbsp;**sheet_name**,  
> &nbsp;&nbsp;&nbsp;&nbsp;**startrow**=0,  
> &nbsp;&nbsp;&nbsp;&nbsp;**startcol**=0,  
> &nbsp;&nbsp;&nbsp;&nbsp;**index**=False,  
> &nbsp;&nbsp;&nbsp;&nbsp;**header**=True,  
> &nbsp;&nbsp;&nbsp;&nbsp;**style**={},  
> &nbsp;&nbsp;&nbsp;&nbsp;**use_base_style**=True,  
> &nbsp;&nbsp;&nbsp;&nbsp;**col_widths**={},  
> &nbsp;&nbsp;&nbsp;&nbsp;**col_autofit**=True,  
> &nbsp;&nbsp;&nbsp;&nbsp;**auto_number_formatting**=True  
> )

> ### Examples:
> ```python
> from beautifulexcel import ExcelWriter
> 
> with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as writer:
>     ws1 = writer.to_excel(
>         example_df,
>         style={"RoE": "bg_light_blue", "D:E": {"fill": "FFEEB7"}},
>         sheet_name="My Sheet",
>         startrow=0,
>         startcol=0,
>         index=True,
>         col_widths={"employees": 100},
>     )
> ```