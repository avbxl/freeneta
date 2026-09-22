PYTHON := python
PIP := $(PYTHON) -m pip
PYINSTALLER := $(PYTHON) -m PyInstaller

APP := freeneta
MAIN := src/freeneta.py

DIST := dist
BUILD := build

.PHONY: all run build compile clean clean-cache install format

all: run

# Run the application directly with Python
run:
	$(PYTHON) $(MAIN)

# Build the executable using PyInstaller
build:
	$(PYINSTALLER) --noconfirm --clean \
		--windowed \
		--name $(APP) \
		--icon app.ico \
		$(MAIN)

# Alias for build
compile: build

# Remove PyInstaller output
clean: clean-cache
	-$(PYTHON) -c "import shutil; shutil.rmtree('$(DIST)', ignore_errors=True)"
	-$(PYTHON) -c "import shutil; shutil.rmtree('$(BUILD)', ignore_errors=True)"
	-$(PYTHON) -c "import os; os.remove('$(APP).spec') if os.path.exists('$(APP).spec') else None"

# Remove every __pycache__ directory recursively
clean-cache:
	$(PYTHON) -c "import os, shutil; [shutil.rmtree(os.path.join(root, d)) for root, dirs, files in os.walk('.') for d in dirs if d == '__pycache__']"

# Install PyInstaller and black
install:
	$(PIP) install pyinstaller black

# Format source code with black
format:
	$(PYTHON) -m black src/
