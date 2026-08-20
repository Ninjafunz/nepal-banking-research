param(
    [Parameter(Position=0)]
    [string]$Task = "help"
)

switch ($Task.ToLower()) {
    "install" {
        python3 -m pip install -e ".[dev]"
    }
    "setup" {
        python3 SCRIPTS/setup_templates.py
    }
    "build" {
        python3 SCRIPTS/calculate_ratios.py
        python3 SCRIPTS/calculate_market_shares.py
        python3 SCRIPTS/build_panel.py
        python3 SCRIPTS/generate_descriptives.py
    }
    "test" {
        python3 -m pytest tests/ -v
    }
    "lint" {
        python3 -m ruff check .
    }
    "dashboard" {
        python3 -m streamlit run app.py
    }
    default {
        Write-Host "Usage: .\tasks.ps1 [install | setup | build | test | lint | dashboard]"
    }
}
