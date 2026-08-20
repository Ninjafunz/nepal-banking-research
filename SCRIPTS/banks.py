"""
banks.py — Master registry of Nepal Class A commercial banks.
Loads from bank_registry.csv and config.yaml.
"""

import os
import pandas as pd
from config.config_loader import load_config

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGISTRY_PATH = os.path.join(BASE_DIR, "bank_registry.csv")

CFG = load_config()

# Load registry
if os.path.exists(REGISTRY_PATH):
    _df = pd.read_csv(REGISTRY_PATH)
    BANKS = _df.to_dict(orient="records")
    # Normalize booleans
    for b in BANKS:
        b["listed"] = bool(b.get("is_listed", True))
        b["state_owned"] = bool(b.get("is_state_owned", False))
        b["code"] = b.get("bank_code")
        b["name"] = b.get("bank_name")
else:
    BANKS = [
        {"code": "NBL",    "name": "Nepal Bank Limited",            "listed": True,  "state_owned": True},
        {"code": "RBB",    "name": "Rastriya Banijya Bank",         "listed": False, "state_owned": True},
        {"code": "ADBL",   "name": "Agriculture Development Bank",  "listed": True,  "state_owned": True},
        {"code": "NABIL",  "name": "Nabil Bank",                    "listed": True,  "state_owned": False},
        {"code": "NIMB",   "name": "Nepal Investment Mega Bank",    "listed": True,  "state_owned": False},
        {"code": "SCB",    "name": "Standard Chartered Bank Nepal", "listed": True,  "state_owned": False},
        {"code": "HBL",    "name": "Himalayan Bank",                "listed": True,  "state_owned": False},
        {"code": "SBI",    "name": "Nepal SBI Bank",                "listed": True,  "state_owned": False},
        {"code": "EBL",    "name": "Everest Bank",                  "listed": True,  "state_owned": False},
        {"code": "BOKL",   "name": "Bank of Kathmandu",             "listed": True,  "state_owned": False},
        {"code": "NICA",   "name": "NIC Asia Bank",                 "listed": True,  "state_owned": False},
        {"code": "MBL",    "name": "Machhapuchchhre Bank",          "listed": True,  "state_owned": False},
        {"code": "KBL",    "name": "Kumari Bank",                   "listed": True,  "state_owned": False},
        {"code": "LLBS",   "name": "Laxmi Sunrise Bank",            "listed": True,  "state_owned": False},
        {"code": "CIVIL",  "name": "Civil Bank",                    "listed": True,  "state_owned": False},
        {"code": "CCBL",   "name": "Century Commercial Bank",       "listed": True,  "state_owned": False},
        {"code": "SANIMA", "name": "Sanima Bank",                   "listed": True,  "state_owned": False},
        {"code": "SBL",    "name": "Siddhartha Bank",               "listed": True,  "state_owned": False},
        {"code": "GIBL",   "name": "Global IME Bank",               "listed": True,  "state_owned": False},
        {"code": "PCBL",   "name": "Prime Commercial Bank",         "listed": True,  "state_owned": False},
        {"code": "PRVU",   "name": "Prabhu Bank",                   "listed": True,  "state_owned": False},
        {"code": "CZBIL",  "name": "Citizens Bank International",   "listed": True,  "state_owned": False},
        {"code": "NMB",    "name": "NMB Bank",                      "listed": True,  "state_owned": False},
    ]

_fy_start = CFG.get("fiscal_years", {}).get("start", 2020)
_fy_end   = CFG.get("fiscal_years", {}).get("end", 2025)
_macro_start = CFG.get("macro_years", {}).get("start", 2020)
_macro_end   = CFG.get("macro_years", {}).get("end", 2026)

FISCAL_YEARS = list(range(_fy_start, _fy_end + 1))
MACRO_YEARS  = list(range(_macro_start, _macro_end + 1))

BANK_CODES = [b["code"] for b in BANKS]
BANK_MAP   = {b["code"]: b["name"] for b in BANKS}
