# -*- coding: utf-8 -*-
import datetime
import numpy as np
import pandas as pd
from beautifulexcel import ExcelWriter



def test_util_get_cell_coordinates():
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


    with ExcelWriter("testing.xlsx", mode="r", theme="elegant_blue") as writer:
        ws1 = writer.to_excel(
            example_df,
            sheet_name="Test Sheet 1",
            startrow=0,
            startcol=0,
            index=True,
            header=True,
        )

        assert ws1.startrow == 0
        assert ws1.startcol == 0
        assert ws1.index_depth == 1
        assert ws1.header_depth == 1
        assert ws1.table_width == example_df.shape[1]
        assert ws1.table_height == example_df.shape[0]
        assert ws1.shape == ((1, 1), (10, 8))
        assert ws1.shape_header == ((1, 1), (2, 8))
        assert ws1.shape_index == ((2, 1), (10, 2))
        assert ws1.shape_body == ((2, 2), (10, 8))

        assert (2, 1) == ws1.util_get_cell_coordinates(ref='A2')
        assert (None, 1) == ws1.util_get_cell_coordinates(ref='A')
        assert (2, None) == ws1.util_get_cell_coordinates(ref='2')
        assert (None, 3) == ws1.util_get_cell_coordinates(ref='employees')

        assert ((2, 1), (4, 3)) == ws1.util_get_range_coordinates('A2:C4')
        assert ((None, 1), (None, 3)) == ws1.util_get_range_coordinates('A:C')
        assert ((1, None), (3, None)) == ws1.util_get_range_coordinates('1:3')
        assert ((None, 3), (None, 3)) == ws1.util_get_range_coordinates(ref='employees')
        assert ((None, 3), (None, 5)) == ws1.util_get_range_coordinates(ref='employees:last_contact')

        assert 'A2:C4' == ws1.util_range_ref_from_coordinates(((2, 1), (4, 3)))

        ws2 = writer.to_excel(
            example_df,
            sheet_name="Test Sheet 2",
            startrow=2,
            startcol=3,
            index=False,
            header=False,
           )

        assert (2, 1) == ws2.util_get_cell_coordinates(ref='A2')
        assert (None, 1) == ws2.util_get_cell_coordinates(ref='A')
        assert (2, None) == ws2.util_get_cell_coordinates(ref='2')
        assert (None, 5) == ws2.util_get_cell_coordinates(ref='employees')




