# -*- coding: utf-8 -*-
import os
from typing import Literal, Any
import numpy as np
import pandas as pd
import openpyxl
import yaml
import datetime
import warnings

from utils import flatten_dict, deepen_dict, excel_range_ref_coordinates, get_custom_styles


class ExcelWriter:
    """
    Class for writing DataFrame objects into excel sheets.

    Example:
        Output pandas dataframe quickly with beautiful formatting.

        >>> from beautifulexcel import ExcelWriter
        >>> with ExcelWriter('workbook.xlsx', mode='r', style='elegant_blue') as writer:
        >>>     ws1 = writer.write_df(df1, sheetname='My Sheet', mode='a', startrow=0, startcol=0)
        >>>     ws1 = writer.write_df(df2, sheetname='My Sheet', mode='a', startrow=20, startcol=0)
    """

    def __init__(
        self,
        file: str,
        mode: Literal["replace", "modify"] = "replace",
        if_sheet_exists: Literal["error", "new", "replace", "overlay"] = "replace",
        theme: str = "elegant_blue",
        date_format: str = None,
        datetime_format: str = None,
        engine_kwargs: dict[str, Any] = {},
        **kwargs,
    ):
        """

        Args:
            file (str): Path to xls or xlsx or ods file
            mode (str): If the file already exists you can either "replace" or "modify" it
            if_sheet_exists (str): If a excel sheet already exists raise an "error", create a "new" sheet with a different name, "replace" the existing sheet with the new one, or "overlay" the new contents with the old ones
            theme (str): Excel style name or path to theme yaml file
            date_format (str): Format string for dates written into Excel files (e. g. 'YYYY-MM-DD')
            datetime_format (str): Format string for datetime objects written into Excel files. (e. g. 'YYYY-MM-DD HH:MM:SS')
            engine_kwargs (str): keywords passed though to openpyxl in "replace"-mode: openpyxl.Workbook(**engine_kwargs); "modify"-mode: openpyxl.load_workbook(file, **engine_kwargs)
        """
        self.file = file

        # modify existing file
        if os.path.isfile(file) and mode == "modify":
            self.writer = pd.ExcelWriter(
                file,
                engine="openpyxl",
                mode="a",
                if_sheet_exists=if_sheet_exists,
                date_format=date_format,
                datetime_format=datetime_format,
                engine_kwargs=engine_kwargs,
                **kwargs,
            )
            # self.writer.book = openpyxl.load_workbook(file, **engine_kwargs)
            # self.writer.sheets = {ws.title: ws for ws in self.writer.book.worksheets}
            self.file_mode = "modify"

        # create new file
        else:
            self.writer = pd.ExcelWriter(
                file,
                engine="openpyxl",
                mode="w",
                if_sheet_exists=None,
                date_format=date_format,
                datetime_format=datetime_format,
                engine_kwargs=engine_kwargs,
                **kwargs,
            )
            # self.writer.book = openpyxl.Workbook(**engine_kwargs)
            self.file_mode = "replace"

        # explicitly no theme defined
        if theme is None or len(theme) == 0:
            self.theme = {}

        # theme defined
        else:
            # standard theme from package
            if "." not in theme:
                this_file_path = os.path.dirname(os.path.abspath(__file__))
                theme = os.path.join(this_file_path, "themes", f"{theme}.yml")

            # read in user yml theme file
            try:
                with open(theme, "r") as file:
                    theme = yaml.safe_load(file)
                    theme_converted = {}
                    # flatten after level 2
                    for level1_name, level1 in theme.items():
                        if level1_name not in theme_converted:
                            theme_converted[level1_name] = {}
                        for level2_name, level2 in level1.items():
                            if level2_name not in theme_converted:
                                theme_converted[level1_name][level2_name] = {}
                            theme_converted[level1_name][level2_name] = flatten_dict(level2)
                    self.theme = theme_converted
            except yaml.YAMLError as exc:
                raise Exception(f"Error when reading in theme file from path '{theme}':", exc)

    def __enter__(self):
        if not hasattr(self, "file"):
            raise Exception(
                "Wrong usage! Please run it this way:\n>>> from beautifulexcel import ExcelWriter\n>>> with ExcelWriter('workbook.xlsx', mode='r', theme='elegant_blue') as writer:\n>>>     ws1 = writer.write_df(df1, sheetname='My Sheet'"
            )
        return self

    def save(self):
        self.writer.save()

    def __exit__(self, type, value, traceback):
        # ToDo: add exception handling here
        self.save()
        self.writer = None
        self = None

    def set_cell_style(self, ws, row_num, col_num, style):
        style_deep = deepen_dict(style)
        cell = ws.cell(row=row_num + 1, column=col_num + 1)

        for style_type, kwargs in style_deep.items():
            if style_type.lower() == "font":
                cell.font = openpyxl.styles.Font(**kwargs)
            elif style_type.lower() == "numfmt" or style_type.lower() == "numberformat":
                cell.number_format = kwargs
            elif style_type.lower() == "align" or style_type.lower() == "alignment":
                cell.alignment = openpyxl.styles.Alignment(**kwargs)
            elif style_type.lower() == "fill" or style_type.lower() == "pfill" or style_type.lower() == "patternfill":
                if type(kwargs) is not dict:
                    kwargs = {"patternType": "solid", "fgColor": kwargs}
                cell.fill = openpyxl.styles.PatternFill(**kwargs)
            elif style_type.lower() == "gfill" or style_type.lower() == "gradientfill":
                cell.fill = openpyxl.styles.GradientFill(**kwargs)
            elif style_type.lower() == "border" or style_type.lower() == "borders":
                border_kwargs = {}
                for border_type, border_props in kwargs.items():
                    border_kwargs[border_type] = openpyxl.styles.Side(**border_props)
                cell.border = openpyxl.styles.Border(**border_kwargs)
            elif style_type.lower() == "protection":
                cell.font = openpyxl.styles.Protection(**kwargs)
            else:
                raise Exception(
                    f'Unknown style type "{style_type}". Available style types are: font, numberformat, align, fill, patternfill, gradientfill, borders, and protection.'
                )

        return cell

    def __apply_style(self, ws, style, startrow, startcol, has_index, index, has_header, header, style_ref_warnings):
        """This internal function applies the table cell styling for the to_excel() function"""
        style_base = style.pop("base", {})
        style_head = {**style_base, **style.pop("head", {})}
        style_index = {**style_base, **style.pop("index", {})}
        style_body_base = {**style_base, **style.pop("body", {})}

        index_depth = index.nlevels if has_index else 0
        header_depth = header.nlevels if has_header else 0

        style_body_custom = {}
        for ref, style in style.items():
            style_coordinates = excel_range_ref_coordinates(
                ref=ref,
                header=header,
                offset_horizontal=index_depth + startcol,
                offset_vertical=header_depth + startrow,
            )

            if style_ref_warnings and (style_coordinates[0] is None or style_coordinates[1] is None):
                warnings.warn(
                    f'Styling ref "{ref}" could not be found. Turn off these warnings by adding style_ref_warnings=False to writer.to_excel().'
                )

            # all good - valid cell style range
            else:
                style_body_custom[style_coordinates] = {**style_body_base, **style}

        table_width = len(header) + index_depth
        table_height = len(index) + header_depth

        # Apply table heading styling
        if has_header:
            for row_num in range(header_depth):
                for col_num in range(table_width):
                    cust_style = get_custom_styles(
                        cell_row_num=row_num,
                        cell_col_num=col_num,
                        styles_custom=style_body_custom,
                        ignore_entire_rows_or_cols=True,
                    )
                    self.set_cell_style(ws=ws, row_num=row_num, col_num=col_num, style={**style_head, **cust_style})

        # Apply table index styling
        if has_index:
            for row_num in range(header_depth, table_height):
                for col_num in range(index_depth):
                    cust_style = get_custom_styles(
                        cell_row_num=row_num,
                        cell_col_num=col_num,
                        styles_custom=style_body_custom,
                        ignore_entire_rows_or_cols=True,
                    )
                    self.set_cell_style(ws=ws, row_num=row_num, col_num=col_num, style={**style_index, **cust_style})

        # Apply table body styling
        for row_num in range(header_depth, table_height):
            for col_num in range(index_depth, table_width):
                cust_style = get_custom_styles(
                    cell_row_num=row_num, cell_col_num=col_num, styles_custom=style_body_custom
                )
                self.set_cell_style(ws=ws, row_num=row_num, col_num=col_num, style={**style_body_base, **cust_style})

    def to_excel(
        self,
        df,
        sheet_name,
        startrow=0,
        startcol=0,
        index=False,
        header=True,
        style={},
        use_base_style=True,
        col_widths={},
        col_autofit=True,
        auto_number_formatting=True,
        style_ref_warnings=True,
    ):
        df.to_excel(
            self.writer, sheet_name=sheet_name, startrow=startrow, startcol=startcol, index=index, header=header
        )
        ws = self.writer.book[sheet_name]

        preset_style = self.theme.get("preset", {})

        # generate final styling that will apply to the dataframe export
        ws_style = {}
        if use_base_style:
            # merge all stylings
            ws_style["base"] = {}

            # add general/base style from style template
            if "general" in self.theme and "base" in self.theme["general"]:
                ws_style["base"] = {**self.theme["general"]["base"]}

            # add table style from style template
            if "table" in self.theme:
                for level, level_styling in self.theme["table"].items():
                    if level not in ws_style:
                        ws_style[level] = {**level_styling}
                    else:
                        ws_style[level] = {**ws_style[level], **level_styling}

        # add number formatting
        if auto_number_formatting:
            # get all date columns
            for col_name, col_series in df.select_dtypes(include=["datetime", "datetimetz"]).items():
                if "date_fmt_iso" in preset_style:
                    ws_style[col_name] = preset_style["date_fmt_iso"]

            # get all numeric columns
            for col_name, col_series in df.select_dtypes(include=["number"]).items():
                col_series_mod = col_series.replace(0, np.nan).abs()
                low = col_series_mod.quantile(0.2)
                high = col_series_mod.quantile(0.8)

                # check if percentages
                if high < 2 and "num_fmt_pct" in preset_style:
                    ws_style[col_name] = preset_style["num_fmt_pct"]
                # check if small number
                elif high < 1_000 and "int" not in str(col_series.dtype) and "num_fmt_decimal" in preset_style:
                    ws_style[col_name] = preset_style["num_fmt_decimal"]
                # check if large number in millions
                elif low > 10_000_000 and "num_fmt_mm" in preset_style:
                    ws_style[col_name] = preset_style["num_fmt_mm"]
                # else normal number format
                elif "num_fmt_general" in preset_style:
                    ws_style[col_name] = preset_style["num_fmt_general"]

        # add styling defined in this function
        for ref, ref_style in style.items():
            if ref not in ws_style:
                ws_style[ref] = {}

            # style is already in base dict granularity
            if type(ref_style) is dict:
                ws_style[ref] = {**ws_style[ref], **ref_style}

            # style is in list of preset names and needs to be extended to dict granularity
            elif type(ref_style) is list:
                for i_ref_style in ref_style:
                    ws_style[ref] = {**ws_style[ref], **preset_style.get(i_ref_style, {})}

            # style is single preset name string and needs to be extended to dict granularity
            else:
                ws_style[ref] = {**ws_style[ref], **preset_style.get(ref_style, {})}

        # actually apply the final themes
        self.__apply_style(
            ws,
            style=ws_style,
            startrow=startrow,
            startcol=startcol,
            has_index=index,
            index=df.index,
            has_header=header,
            header=df.columns,
            style_ref_warnings=style_ref_warnings,
        )

        # autofit columns
        if col_autofit:
            CHARACTER_FACTOR = 1.28

            # enumerate though indeices and columns
            for i, col in enumerate(
                ([df.index.get_level_values(i) for i in range(df.index.nlevels)] if index else [])
                + [i for _, i in df.items()],
                start=startcol,
            ):
                # is datetime column
                if pd.api.types.is_datetime64_any_dtype(col):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = int(10 * CHARACTER_FACTOR)
                # is string column
                elif pd.api.types.is_string_dtype(col):
                    try:
                        length = np.quantile(col.str.len().values, 0.8)
                        ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = int(
                            max(length, 4) * CHARACTER_FACTOR
                        )
                    except:
                        # ignore error
                        pass
                # is numeric column
                elif pd.api.types.is_numeric_dtype(col):
                    length = len("{:,}".format(int(col.quantile(0.8))))
                    ws.column_dimensions[openpyxl.utils.get_column_letter(i + 1)].width = int(
                        max(length, 6) * CHARACTER_FACTOR
                    )

            # apply column widths
            for col_ref, col_width in col_widths.items():
                col_coordinates = excel_range_ref_coordinates(
                    ref=col_ref,
                    header=df.columns,
                    offset_horizontal=(df.index.nlevels if index else 0) + startcol,
                    offset_vertical=(df.columns.nlevels if header else 0) + startrow,
                    )

                if col_coordinates[0] is not None and col_coordinates[0][1] is not None and col_coordinates[1] is not None and col_coordinates[1][1] is not None:
                    for col_idx in range(col_coordinates[0][1], col_coordinates[1][1] + 1):
                        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = col_width


        return ws


if __name__ == "__main__":
    example_df = pd.DataFrame(
        {
            "client": ["A", "B", "C", "D", "E", "F", "G", "H"],
            "industry": [
                "ASEET MANAGEMENT",
                "ASEET MANAGEMENT",
                "BANK",
                "INSURANCE",
                "VERY VERY VERY VERY VERY LONG INSUSTRY NAME",
                "BANK",
                "COMMODITY BROKER",
                "INSURANCE",
            ],
            "employees": [25_000, 17_000_000, 14, 9_000, 12_000_000, 9_000, np.nan, 4_000_000],
            "inception": [
                datetime.datetime(2022, 1, 1),
                datetime.datetime(2000, 10, 30),
                datetime.datetime(2010, 5, 6),
                np.nan,
                datetime.datetime(1997, 1, 1),
                datetime.datetime(1962, 1, 1),
                datetime.datetime(2003, 11, 10),
                datetime.datetime(2022, 1, 1),
            ],
            "last_contact": [
                datetime.datetime(2022, 1, 1, 10, 2, 5),
                datetime.datetime(2000, 10, 30),
                datetime.datetime(2010, 5, 6),
                np.nan,
                datetime.datetime(1997, 1, 1),
                datetime.datetime(1962, 1, 1),
                datetime.datetime(2003, 11, 10, 10, 2, 5),
                datetime.datetime(2022, 1, 1),
            ],
            "RoE": [0.05, -0.05, 0.15, 1.05, -0.02, np.nan, 0.08, 0.05],
            "revenue": [
                50_000_000.4387,
                np.nan,
                63_000_000.4387,
                25_000.4387,
                25_000_000.4387,
                -50_000_000.4387,
                76_000_000.4387,
                25_000.4387,
            ],
        }
    )
    example_df.set_index("client", inplace=True)

    with ExcelWriter("workbook.xlsx", mode="r", theme="elegant_blue") as writer:
        ws1 = writer.to_excel(
            example_df,
            style={"RoE": "bg_light_blue", "D:E": {"fill": "FFEEB7"}},
            sheet_name="My Sheet",
            startrow=0,
            startcol=0,
            index=True,
            col_widths={'employees': 100}
        )

    # example_df.to_excel('test.xlsx', sheet_name='My Sheet', startrow=0, startcol=0, index=True)
