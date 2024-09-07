from typing import Literal
import pandas as pd


class ExcelWriter:
    """
    Class for writing DataFrame objects into excel sheets.

    Example:
        Output pandas dataframe quickly with beautiful fromatting.

        >>> from beautifulexcel import ExcelWriter
        >>> with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as wb:
        >>>     ws1 = wb.write_df(df1, sheetname='My Sheet', mode='a', startrow=0, startcol=0)
        >>>     ws1 = wb.write_df(df2, sheetname='My Sheet', mode='a', startrow=20, startcol=0)
    """
    def __init__(self, file: str, mode: Literal['r', 'm'] = "r", theme: str ='elegant_blue'):
        """

        Args:
            file (str): Path to xls or xlsx or ods file
            mode (str): If file already exists "r" for replace file or "m" for modify file
            theme (str): Excel theme name or path to theme yaml file
        """
        print('Test Successful')


#writer = pd.ExcelWriter("test.xlsx", engine="xlsxwriter")