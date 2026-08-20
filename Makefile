.PHONY: help install setup build test lint clean dashboard

help:
	@echo "Nepal Banking Research Platform — Make Commands"
	@echo "  make install    - Install package dependencies"
	@echo "  make setup      - Generate blank template workbooks"
	@echo "  make build      - Run full calculation and build multi-format master panel"
	@echo "  make test       - Execute pytest test suite"
	@echo "  make lint       - Check code style with ruff"
	@echo "  make clean      - Clean cache and temporary files"
	@echo "  make dashboard  - Launch interactive Streamlit web app"

install:
	pip install -e ".[dev]"

setup:
	python SCRIPTS/setup_templates.py

build:
	python SCRIPTS/calculate_ratios.py
	python SCRIPTS/calculate_market_shares.py
	python SCRIPTS/build_panel.py
	python SCRIPTS/generate_descriptives.py

test:
	pytest tests/ -v

lint:
	ruff check .

clean:
	rm -rf __pycache__ .pytest_cache SCRIPTS/__pycache__ tests/__pycache__

dashboard:
	streamlit run app.py
