# Define paths (adjust as needed)
VENV = .venv
PYTHON = $(VENV)/bin/python
PIP = $(VENV)/bin/pip

SITE = https://excel-api.fly.dev
LOCAL_SERVER = http://localhost:1212
PORT = 1212

.DEFAULT_GOAL := run
.ONESHELL:

#PHONY targets to avoid conflicts with existing files/directories
.PHONY: all help black clean venv

all: help

help:
		@echo "Available commands:"
		@echo "  venv                   - sourc env and install dep"
		@echo "  run-local              - run with local server"
		@echo "  run                    - run with fly server"
		@echo "  black                  - Format code with Black"
		@echo "  clean                  - Remove compiled Python files (.pyc)"

$(VENV)/bin/activate:  requirements.txt
		@echo "create env ..."
		@python -m venv .venv

venv: $(VENV)/bin/activate
		@echo "source venv ..."
		@. $(VENV)/bin/activate
		@$(PIP) install -r  requirements.txt

run-local: venv
		@echo "run localy ..."
		@export SERVER_URL=$(LOCAL_SERVER)
		@$(PYTHON) ./main.py

run-server: venv
		@echo "run with server ..."
		@export SERVER_URL=$(SITE)
		@python ./main.py

run:run-server

black:
		@echo "format code ..."
		@black $(PROJECT_DIR)

clean:
		@echo "clean directories ..."
		@find . -name "*.pyc" -delete

