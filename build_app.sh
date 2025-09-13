#!/bin/bash

# Get the absolute path of the project root directory (the directory where the script is located)
PROJECT_ROOT="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
VENV_DIR="$PROJECT_ROOT/build_env"
REQUIREMENTS_FILE="$PROJECT_ROOT/requirements.txt"

echo "Project Root: $PROJECT_ROOT"
echo "Virtual Environment: $VENV_DIR"

# --- CRITICAL: Force remove the old virtual environment to ensure a clean slate ---
echo "Removing old virtual environment to ensure a clean build..."
rm -rf "$VENV_DIR"

# --- New Step: Attempt to fix broken tkinter with Homebrew ---
if command -v brew &> /dev/null; then
    echo "Homebrew detected. Attempting to fix tkinter installation..."
    # This command is crucial to fix potential issues with Homebrew's Python and tkinter linkage.
    brew reinstall python-tk@3.9
    if [ $? -ne 0 ]; then
        echo "Warning: 'brew reinstall python-tk@3.9' failed. Continuing build, but GUI may fail if tkinter is broken."
    else
        echo "tkinter fix attempt completed."
    fi
else
    echo "Homebrew not found. Skipping tkinter fix."
fi

# --- Step 1: Create and activate the virtual environment ---
if [ ! -d "$VENV_DIR" ]; then
    echo "Creating virtual environment with Python 3.9..."
    python3.9 -m venv "$VENV_DIR"
    if [ $? -ne 0 ]; then
        echo "Error: Failed to create virtual environment."
        exit 1
    fi
fi

echo "Activating virtual environment..."
source "$VENV_DIR/bin/activate"

# --- Step 2: Install dependencies in the virtual environment ---
echo "Installing dependencies from $REQUIREMENTS_FILE..."
pip install --upgrade pip
pip install -r "$REQUIREMENTS_FILE"
if [ $? -ne 0 ]; then
    echo "Error: Failed to install dependencies."
    deactivate
    exit 1
fi

# --- Step 3: Run PyInstaller in the virtual environment ---
echo "Starting the packaging process with final compatibility settings..."

# Set the deployment target to ensure compatibility with older macOS versions.
export MACOSX_DEPLOYMENT_TARGET=10.13

pyinstaller \
    --name "ThroughputAnalyzer" \
    --onedir \
    --windowed \
    --distpath "$PROJECT_ROOT/dist" \
    --workpath "$PROJECT_ROOT/build" \
    --clean \
    --noupx \
    --hidden-import="tkinter" \
    --hidden-import="pandas" \
    --hidden-import="matplotlib" \
    --hidden-import="tkinterdnd2" \
    --exclude-module "markupsafe._speedups" \
    --target-arch x86_64 \
    "$PROJECT_ROOT/APP/target.py"

# Check the build result
if [ $? -eq 0 ]; then
    echo "Build successful!"
    echo "The application is located in: $PROJECT_ROOT/dist"
else
    echo "Build failed. Please check the output above for errors."
fi

# Unset the environment variable
unset MACOSX_DEPLOYMENT_TARGET

# --- Step 4: Deactivate the virtual environment ---
echo "Deactivating virtual environment."
deactivate

exit 0