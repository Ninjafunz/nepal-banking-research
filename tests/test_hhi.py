import pytest
import pandas as pd
from SCRIPTS.calculate_market_shares import compute_hhi

def test_hhi_equal_shares():
    # 4 firms with 25% share each -> HHI = 4 * 25^2 = 2500
    shares = pd.Series([25.0, 25.0, 25.0, 25.0])
    assert compute_hhi(shares, "percent_squared") == 2500.0
    assert pytest.approx(compute_hhi(shares, "decimal"), 0.0001) == 0.25

def test_hhi_monopoly():
    # Single firm with 100% share -> HHI = 10,000
    shares = pd.Series([100.0])
    assert compute_hhi(shares, "percent_squared") == 10000.0

def test_hhi_empty():
    assert compute_hhi(pd.Series([]), "percent_squared") == 0.0
