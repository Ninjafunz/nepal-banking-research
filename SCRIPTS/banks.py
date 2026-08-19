"""
banks.py - Master list of Nepal Class A commercial banks (NRB licensed).
Adjust this list as mergers occur or new banks are licensed.
All codes follow NRB / NEPSE ticker conventions.
"""

BANKS = [
    {"code": "NBL",    "name": "Nepal Bank Limited",            "listed": True,  "state_owned": True},
    {"code": "RBB",    "name": "Rastriya Banijya Bank",         "listed": True,  "state_owned": True},
    {"code": "ADBL",   "name": "Agriculture Development Bank",  "listed": True,  "state_owned": True},
    {"code": "NABIL",  "name": "Nabil Bank",                    "listed": True,  "state_owned": False},
    {"code": "NIMB",   "name": "Nepal Investment Mega Bank",    "listed": True,  "state_owned": False},
    {"code": "SCB",    "name": "Standard Chartered Bank Nepal", "listed": True,  "state_owned": False},
    {"code": "HBL",    "name": "Himalayan Bank",                "listed": True,  "state_owned": False},
    {"code": "SBI",    "name": "Nepal SBI Bank",                "listed": True,  "state_owned": False},
    {"code": "EBL",    "name": "Everest Bank",                  "listed": True,  "state_owned": False},
    {"code": "BOKL",   "name": "Bank of Kathmandu Lumbini",     "listed": True,  "state_owned": False},
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

FISCAL_YEARS = list(range(2020, 2026))   # FY2020 to FY2025
MACRO_YEARS  = list(range(2020, 2027))   # FY2020 to FY2026 for macro/market datasets

BANK_CODES = [b["code"] for b in BANKS]
BANK_MAP   = {b["code"]: b["name"] for b in BANKS}
