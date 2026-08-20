import pandas as pd

from SCRIPTS.validators import validate_balance_sheet, validate_ratios


def test_validate_balance_sheet_correct():
    df = pd.DataFrame([{
        "bank_code": "TEST", "fy": 2024,
        "total_assets": 100.0, "total_liabilities": 88.0, "shareholders_equity": 12.0
    }])
    is_valid, flags = validate_balance_sheet(df)
    assert is_valid is True
    assert len(flags) == 0

def test_validate_balance_sheet_equity_exceeds_assets():
    df = pd.DataFrame([{
        "bank_code": "TEST", "fy": 2024,
        "total_assets": 100.0, "total_liabilities": 20.0, "shareholders_equity": 120.0
    }])
    is_valid, flags = validate_balance_sheet(df)
    assert is_valid is False
    assert any("equity" in f and ">= total_assets" in f for f in flags)

def test_validate_ratios_car_breach():
    df = pd.DataFrame([{
        "bank_code": "TEST", "fy": 2024,
        "roa": 1.5, "roe": 12.0, "nim": 3.8, "car_pct": 7.2
    }])
    is_valid, flags = validate_ratios(df)
    assert is_valid is False
    assert any("REGULATORY BREACH" in f for f in flags)
