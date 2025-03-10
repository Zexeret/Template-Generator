# Document Generator

This project reads data from a `.xlsx` or `.csv` file and replaces placeholders in a `.docx` template using configurations from `.json` files.

## ⚡ Features
- Reads data from `inputPath` path set in config.
- Uses `config/*.json` to map placeholders
- Replaces placeholders in a `.docx` template
- Saves the modified document with updated values

---

## 🚀 Setup Guide

### 🔹 **Clone the Repository**
```sh
$ git clone https://github.com/your-repo/python-doc-generator.git

$ cd python-doc-generator
```

### 🔹 **Create and Activate Virtual Environment**
```sh
$ python -m venv tsgen

$ tsgen\Scripts\activate
```

### 🔹 **Install Dependencies**
```sh
$ pip install -r requirements.txt
```

### 🔹 **Run Script**
Make sure you have your virtual environment activated, if not you should run the first command else skip it.

```sh
$ tsgen\Scripts\activate
$ python ts_script.py
```

### 🔹 **Build Executable File**
Make sure you have your virtual environment activated, if not you should run the first command else skip it.
It copies the content of your current product folder.

```sh
$ tsgen\Scripts\activate
$ python build.py
```

For details about various arguments:
> python ts_script.py --help

Example: To run in verbose mode
> python ts_script.py -v