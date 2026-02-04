# SlipDayCounter Makefile
#
# Targets:
#   make run      Ensure venv + deps, then launch Bootstrap (splash -> app)
#   make install  Ensure venv + deps only
#   make clean    Remove venv + caches
#   make doctor   Verify tkinter + pandas in the venv

.SILENT:

# Prefer a Python that includes Tkinter on macOS. Homebrew python can lack _tkinter.
PYTHON ?= /usr/local/bin/python3
ifeq (,$(wildcard $(PYTHON)))
PYTHON := python3
endif

VENV_DIR ?= .venv
VENV_PY := $(VENV_DIR)/bin/python
VENV_PIP := $(VENV_PY) -m pip

# Quieter pip output (hides "Requirement already satisfied" noise)
PIP_FLAGS ?= -q --disable-pip-version-check --no-input

ARGS ?=
IMPORT_DIR ?= import

SCRIPT ?= CountSlips.py
BOOTSTRAP ?= Bootstrap.py

REQUIREMENTS ?= requirements.txt
DEPS ?= pandas

.PHONY: run install clean doctor venv run_inner

# Create venv if missing
$(VENV_PY):
	$(PYTHON) -m venv $(VENV_DIR)
	$(VENV_PIP) install $(PIP_FLAGS) --upgrade pip setuptools wheel

venv: $(VENV_PY)

install: $(VENV_PY)
	@if [ -f "$(REQUIREMENTS)" ]; then \
		printf "Installing from %s\n" "$(REQUIREMENTS)"; \
		$(VENV_PIP) install $(PIP_FLAGS) -r "$(REQUIREMENTS)"; \
		$(VENV_PIP) install $(PIP_FLAGS) $(DEPS); \
	else \
		printf "Installing deps: %s\n" "$(DEPS)"; \
		$(VENV_PIP) install $(PIP_FLAGS) $(DEPS); \
	fi

# Run script directly (no splash)
run_inner: install
	SLIPDAYCOUNTER_SKIP_MAKE=1 $(VENV_PY) $(SCRIPT) --import-dir $(IMPORT_DIR) $(ARGS)

# Canonical launcher (splash -> app)
run: install
	ARGS="$(ARGS)" IMPORT_DIR="$(IMPORT_DIR)" SCRIPT="$(SCRIPT)" VENV_DIR="$(VENV_DIR)" PYTHON="$(PYTHON)" SLIPDAYCOUNTER_SKIP_MAKE=1 $(VENV_PY) $(BOOTSTRAP)

doctor: install
	$(VENV_PY) -c "import tkinter; print('tk ok')"
	$(VENV_PY) -c "import pandas; print('pandas ok')"

clean:
	rm -rf $(VENV_DIR)
	rm -rf __pycache__ .pytest_cache .mypy_cache
	find . -name "*.pyc" -delete