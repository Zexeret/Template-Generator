import os
import shutil
import subprocess

# Creates an exe file inside TOOL/TS_Generator folder
# Copies whole product folder inside TOOL/TS_Generator folder


# Get the base directory dynamically
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Define necessary paths
BUILD_DIR = os.path.join(BASE_DIR, "TOOL")
TS_GEN_DIR = os.path.join(BUILD_DIR, "TS_Generator")
PRODUCTS_DIR = os.path.join(BASE_DIR, "products")
DIST_DIR = os.path.join(BASE_DIR, "dist")
EXE_NAME = "ts-generator.exe"

# Clear TOOL directory if it exists
if os.path.exists(BUILD_DIR):
    shutil.rmtree(BUILD_DIR)


# Create TS_Generator directory inside TOOL
os.makedirs(TS_GEN_DIR, exist_ok=True)

# Copy the entire products folder to TS_Generator
shutil.copytree(PRODUCTS_DIR, os.path.join(TS_GEN_DIR, "products"), dirs_exist_ok=True)

# Run PyInstaller command
pyinstaller_cmd = [
    "pyinstaller",
    "--onefile",
    "--add-data", f"{os.path.join(BASE_DIR, 'config')};config",
    "--add-data", f"{os.path.join(BASE_DIR, 'customUtil.py')};.",
    "--name", "ts-generator",
    os.path.join(BASE_DIR, "ts_script.py"),
]

subprocess.run(pyinstaller_cmd, shell=True, check=True)

# Copy the generated executable into TS_Generator
EXE_PATH = os.path.join(DIST_DIR, EXE_NAME)
if os.path.exists(EXE_PATH):
    shutil.copy(EXE_PATH, TS_GEN_DIR)
    print(f"Executable copied to {TS_GEN_DIR}")
else:
    print("Error: Executable not found in dist folder!")

print("Process completed successfully!")
