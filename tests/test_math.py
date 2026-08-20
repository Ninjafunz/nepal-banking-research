import pytest
import pandas as pd
import numpy as np
from SCRIPTS.calculate_ratios import safe_div, avg_balance, yoy_growth

def test_safe_div_zero_and_nan():
    a = pd.Series([10.0, 20.0, 30.0])
    b = pd.Series([2.0, 0.0, np.nan])
    res = safe_div(a, b)
    assert res[0] == 5.0
    assert np.isnan(res[1])
    assert np.isnan(res[2])

def test_avg_balance_calculation():
    df = pd.DataFrame({
        "bank_code": ["B1", "B1", "B1"],
        "fy": [2022, 2023, 2024],
        "assets": [100.0, 120.0, 150.0]
    })
    res_df = avg_balance(df, "assets", key="bank_code")
    assert np.isnan(res_df["avg_assets"].iloc[0])
    assert res_df["avg_assets"].iloc[1] == 110.0
    assert res_df["avg_assets"].iloc[2] == 135.0

def test_yoy_growth():
    df = pd.DataFrame({
        "bank_code": ["B1", "B1"],
        "fy": [2023, 2024],
        "loans": [100.0, 115.0]
    })
    res_df = yoy_growth(df, "loans", key="bank_code")
    assert np.isnan(res_df["loans_growth"].iloc[0])
    assert pytest.approx(res_df["loans_growth"].iloc[1], 0.01) == 15.0
