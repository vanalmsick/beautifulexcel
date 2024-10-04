# -*- coding: utf-8 -*-
import re
import pandas as pd


def dict_extend_with_dict(dict_obj, key, value_dict):
    if key not in dict_obj:
        dict_obj[key] = {**value_dict}
    else:
        dict_obj[key] = {**dict_obj[key], **value_dict}
    return dict_obj


def flatten_dict(dict_obj: dict, sep="__"):
    """Flatten a multi-level dictionary e.g. {'a1': {'b': 1}, 'a2': 3} -> {'a1__b': 1, 'a2': 3}"""
    df = pd.json_normalize(dict_obj, sep=sep)
    d_flat = df.to_dict(orient="records")[0]
    return {k[2:] if len(k) > 2 and k[:2] == '__' else k: v for k, v in d_flat.items()}


def deepen_dict(dict_obj, sep="__"):
    """The opposite of flatten_dict e.g. {'a1__b': 1, 'a2': 3} -> {'a1': {'b': 1}, 'a2': 3}"""
    parsed_row = {}
    for col_label, v in dict_obj.items():
        keys = col_label.split(sep)

        current = parsed_row
        for i, k in enumerate(keys):
            if i == len(keys) - 1:
                current[k] = v
            else:
                if k not in current.keys():
                    current[k] = {}
                current = current[k]
    return parsed_row


def excel_column_name(n):
    """Number to Excel-style column name, e.g., 1 = A, 26 = Z, 27 = AA, 703 = AAA."""
    name = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        name = chr(r + ord("A")) + name
    return name


def excel_column_number(name):
    """Excel-style column name to number, e.g., A = 1, Z = 26, AA = 27, AAA = 703."""
    n = 0
    for c in name:
        n = n * 26 + 1 + ord(c) - ord("A")
    return n


def is_valid_excel_cell(cell):
    """Check if valid single excel cell ref like 'B2', 'AB100' etc."""
    m = re.match(r"^([A-Z]{1,2}|[A-W][A-Z]{2}|X[A-E][A-Z]|XF[A-D])([1-9]\d{0,6})$", cell)
    valid = bool(m) and int(m.group(2)) < 1_048_577
    return (valid, (m.group(1) if valid else None), (int(m.group(2)) if valid else None))


def excel_cell_ref_coordinates(ref, header=None, offset_horizontal=0, offset_vertical=0):
    """Transform references like 'A', 'COLUMN_A', '2', 'C2' to (row_num, col_num)"""
    # if already tuple - nothing to do
    if type(ref) is tuple:
        return ref

    # check if (relative) row number
    if ref.isdigit():
        return (int(ref) - 1 + offset_vertical, None)

    # check if column name
    if header is not None and ref in header:
        return (None, header.get_loc(ref) + offset_horizontal)

    # check if (reasonable) column letter
    if len(ref) < 4 and ref.isalpha() and excel_column_number(ref) < 1_048_577:
        return (None, excel_column_number(ref) - 1)

    # check if valid excel cell ref
    valid_excel_cell, col_letter, row_num = is_valid_excel_cell(ref)
    if valid_excel_cell:
        return (row_num - 1, excel_column_number(col_letter) - 1)


def excel_range_ref_coordinates(ref, header, offset_horizontal=0, offset_vertical=0):
    """Transform references like 'A:C', 'COLUMN_A:COLUMN_C', '2:3', 'C2:D4' to ((row_num, col_num), (row_num, col_num))"""
    if type(ref) is tuple:
        ref_start, ref_end = ref

    elif type(ref) is str:
        # ref is range
        if ":" in ref:
            ref_start, ref_end = ref.split(":")
        # ref is no range
        else:
            ref_start = ref_end = ref

    ref_start = excel_cell_ref_coordinates(
        ref=ref_start, header=header, offset_horizontal=offset_horizontal, offset_vertical=offset_vertical
    )
    ref_end = excel_cell_ref_coordinates(
        ref=ref_end, header=header, offset_horizontal=offset_horizontal, offset_vertical=offset_vertical
    )

    if ref_start is None or ref_end is None:
        # cell range is invalid
        return (ref_start, ref_end)
    else:
        # ensure valid ascending cell ref - e.g. C1:A5 -> A1:C5
        row_start, col_start = ref_start
        row_end, col_end = ref_end

        if row_start is not None and row_end is not None and row_end < row_start:
            row_start, row_end = row_end, row_start
        if col_start is not None and col_end is not None and col_end < col_start:
            col_start, col_end = col_end, col_start

        return ((row_start, col_start), (row_end, col_end))


def cell_within_cell_range(cell_row_num, cell_col_num, cell_range, ignore_entire_rows_or_cols=False):
    """Check if a cell e.g. 'A1' (0, 0) is within a cell range e.g. ((0, 0), (1, 1))"""
    ((range_start_row, range_start_col), (range_end_row, range_end_col)) = cell_range

    # if ignore_entire_rows_or_cols is True
    if ignore_entire_rows_or_cols and any(
        [range_start_row is None, range_start_col is None, range_end_row is None, range_end_col is None]
    ):
        return False

    # outside on the top
    if range_start_row is not None and range_start_row > cell_row_num:
        return False

    # outside on the bottom
    if range_end_row is not None and range_end_row < cell_row_num:
        return False

    # outside on the left
    if range_start_col is not None and range_start_col > cell_col_num:
        return False

    # outside on the right
    if range_end_col is not None and range_end_col < cell_col_num:
        return False

    # inside
    return True


def get_custom_styles(cell_row_num, cell_col_num, styles_custom, ignore_entire_rows_or_cols=False):
    """Check if custom style needs to be applied to current cell and return custom style"""
    cell_style = {}
    for cell_range, cust_style in styles_custom.items():
        # if current cell within cell range of cutsom style add style
        if cell_within_cell_range(
            cell_row_num=cell_row_num,
            cell_col_num=cell_col_num,
            cell_range=cell_range,
            ignore_entire_rows_or_cols=ignore_entire_rows_or_cols,
        ):
            cell_style = {**cell_style, **cust_style}

    return cell_style
