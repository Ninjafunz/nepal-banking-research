import os
import yaml

DEFAULT_CONFIG = {
    "project_name": "nepal-banking-research",
    "version": "1.0.0",
    "currency_unit": "NPR Millions",
    "average_balance_method": True,
    "fiscal_years": {"start": 2020, "end": 2025},
    "macro_years": {"start": 2020, "end": 2026},
    "hhi_convention": "percent_squared",
    "hhi_thresholds": {"unconcentrated_max": 1000, "moderately_concentrated_max": 1800},
    "export_formats": ["xlsx", "csv", "parquet", "dta"],
    "validation_rules": {
        "roa_min": -10.0,
        "roa_max": 15.0,
        "roe_min": -50.0,
        "roe_max": 60.0,
        "nim_min": 0.0,
        "nim_max": 12.0,
        "car_min_regulatory": 8.5,
        "cost_income_min": 5.0,
        "cost_income_max": 150.0,
        "balance_sheet_tolerance_pct": 1.0
    }
}

def load_config(config_path=None):
    if config_path is None:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        config_path = os.path.join(base_dir, "config.yaml")
    
    if not os.path.exists(config_path):
        return DEFAULT_CONFIG
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
            return cfg or DEFAULT_CONFIG
    except Exception:
        return DEFAULT_CONFIG
