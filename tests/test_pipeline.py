import os
import pytest
import pandas as pd
from tests.generate_mock_data import generate_synthetic_fixture
from SCRIPTS.build_panel import build

def test_end_to_end_panel_build(tmp_path):
    # Verify building panel executes cleanly and outputs valid dataframe
    panel = build()
    assert isinstance(panel, pd.DataFrame)
    assert len(panel) > 0
    assert "total_assets" in panel.columns
    assert "roa" in panel.columns
