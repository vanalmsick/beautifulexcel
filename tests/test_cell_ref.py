# -*- coding: utf-8 -*-
import pandas as pd
from beautifulexcel.utils import excel_cell_ref_coordinates, excel_range_ref_coordinates, cell_within_cell_range


def test_cell_single_level_df():
    df = pd.DataFrame({"calories": [420, 380, 390, 210], "duration": [50, 40, 45, 90], "test": [4, 1, 4, 7]})

    assert (None, 1) == excel_cell_ref_coordinates("duration", df.columns)
    assert (None, 1) == excel_cell_ref_coordinates("B", df.columns)
    assert (0, None) == excel_cell_ref_coordinates("1", df.columns)
    assert (1, 1) == excel_cell_ref_coordinates("B2", df.columns)
    assert (2, 2) == excel_cell_ref_coordinates((2, 2), df.columns)


def test_range_single_level_df():
    df = pd.DataFrame({"calories": [420, 380, 390, 210], "duration": [50, 40, 45, 90], "test": [4, 1, 4, 7]})

    assert ((None, 1), (None, 2)) == excel_range_ref_coordinates(
        "duration:test", df.columns, offset_horizontal=0, offset_vertical=0
    )
    assert ((None, 1), (None, 2)) == excel_range_ref_coordinates(
        "duration:test", df.columns, offset_horizontal=0, offset_vertical=1
    )
    assert ((None, 2), (None, 3)) == excel_range_ref_coordinates(
        "duration:test", df.columns, offset_horizontal=1, offset_vertical=0
    )
    assert ((0, None), (2, None)) == excel_range_ref_coordinates(
        "1:3", df.columns, offset_horizontal=0, offset_vertical=0
    )
    assert ((1, None), (3, None)) == excel_range_ref_coordinates(
        "1:3", df.columns, offset_horizontal=0, offset_vertical=1
    )
    assert ((0, None), (2, None)) == excel_range_ref_coordinates(
        "1:3", df.columns, offset_horizontal=1, offset_vertical=0
    )
    assert ((None, 1), (None, 2)) == excel_range_ref_coordinates(
        "B:C", df.columns, offset_horizontal=5, offset_vertical=5
    )
    assert ((0, 0), (1, 2)) == excel_range_ref_coordinates("A1:C2", df.columns, offset_horizontal=5, offset_vertical=5)
    assert ((1, 1), (5, 6)) == excel_range_ref_coordinates(((5, 1), (1, 6)), df.columns)
    assert ((0, 0), (1, 2)) == excel_range_ref_coordinates("A2:C1", df.columns, offset_horizontal=5, offset_vertical=5)
    assert ((0, 0), (1, 2)) == excel_range_ref_coordinates("C1:A2", df.columns, offset_horizontal=5, offset_vertical=5)


def test_cell_within_cell_range():
    cell_range = excel_range_ref_coordinates("B2:D4", None)

    a1 = excel_cell_ref_coordinates("A1", None)
    b1 = excel_cell_ref_coordinates("B1", None)
    c1 = excel_cell_ref_coordinates("C1", None)
    d1 = excel_cell_ref_coordinates("D1", None)
    e1 = excel_cell_ref_coordinates("E1", None)

    assert ~cell_within_cell_range(a1[0], a1[1], cell_range)
    assert ~cell_within_cell_range(b1[0], a1[1], cell_range)
    assert ~cell_within_cell_range(c1[0], a1[1], cell_range)
    assert ~cell_within_cell_range(d1[0], a1[1], cell_range)
    assert ~cell_within_cell_range(e1[0], e1[1], cell_range)

    b2 = excel_cell_ref_coordinates("B2", None)
    c3 = excel_cell_ref_coordinates("C3", None)
    d4 = excel_cell_ref_coordinates("D4", None)
    c2 = excel_cell_ref_coordinates("C2", None)
    d2 = excel_cell_ref_coordinates("D2", None)
    d3 = excel_cell_ref_coordinates("D3", None)
    b4 = excel_cell_ref_coordinates("B4", None)

    assert cell_within_cell_range(b2[0], b2[1], cell_range)
    assert cell_within_cell_range(c3[0], c3[1], cell_range)
    assert cell_within_cell_range(d4[0], d4[1], cell_range)
    assert cell_within_cell_range(c2[0], c2[1], cell_range)
    assert cell_within_cell_range(d2[0], d2[1], cell_range)
    assert cell_within_cell_range(d3[0], d3[1], cell_range)
    assert cell_within_cell_range(b4[0], b4[1], cell_range)

    a3 = excel_cell_ref_coordinates("A3", None)
    a4 = excel_cell_ref_coordinates("A4", None)
    a5 = excel_cell_ref_coordinates("A5", None)
    c5 = excel_cell_ref_coordinates("C5", None)
    d5 = excel_cell_ref_coordinates("D5", None)
    e5 = excel_cell_ref_coordinates("E5", None)
    e4 = excel_cell_ref_coordinates("E4", None)
    e3 = excel_cell_ref_coordinates("E3", None)

    assert ~cell_within_cell_range(a3[0], a3[1], cell_range)
    assert ~cell_within_cell_range(a4[0], a4[1], cell_range)
    assert ~cell_within_cell_range(a5[0], a5[1], cell_range)
    assert ~cell_within_cell_range(c5[0], c5[1], cell_range)
    assert ~cell_within_cell_range(d5[0], d5[1], cell_range)
    assert ~cell_within_cell_range(e5[0], e5[1], cell_range)
    assert ~cell_within_cell_range(e4[0], e4[1], cell_range)
    assert ~cell_within_cell_range(e3[0], e3[1], cell_range)

    col_range = excel_range_ref_coordinates("B:C", None)
    row_range = excel_range_ref_coordinates("2:3", None)

    a1 = excel_cell_ref_coordinates("A1", None)
    b2 = excel_cell_ref_coordinates("B2", None)
    c3 = excel_cell_ref_coordinates("C3", None)
    d4 = excel_cell_ref_coordinates("D4", None)

    assert ~cell_within_cell_range(a1[0], a1[1], col_range)
    assert cell_within_cell_range(b2[0], b2[1], col_range)
    assert cell_within_cell_range(c3[0], c3[1], col_range)
    assert ~cell_within_cell_range(d4[0], d4[1], col_range)

    assert ~cell_within_cell_range(a1[0], a1[1], row_range)
    assert cell_within_cell_range(b2[0], b2[1], row_range)
    assert cell_within_cell_range(c3[0], c3[1], row_range)
    assert ~cell_within_cell_range(d4[0], d4[1], row_range)
