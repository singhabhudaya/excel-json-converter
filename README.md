# Excel-JSON Converter

A small Tkinter-based Python GUI (bundled as a standalone Windows EXE via PyInstaller) that converts `.xlsx`/`.xls` files into JSON. No Python installation is required.

---

## Table of Contents

1. [Overview](#overview)  
2. [Features](#features)  
3. [Installation](#installation)  
4. [Usage](#usage)  
5. [Download](#download)  
6. [License](#license)  

---

## Overview

When you have data in an Excel spreadsheet and need a quick JSON export, this tool does the heavy lifting for you. Simply run the EXE, pick your `.xlsx` or `.xls` file, and it spits out a prettified JSON array where each sheet row becomes a JSON object.

- **Built with:**  
  - Python 3  
  - pandas  
  - openpyxl  
  - Tkinter (for GUI)  
  - PyInstaller (to bundle into a single `.exe`)  

---

## Features

- **Standalone Windows EXE**: Runs without requiring any Python installation on the target machine.  
- **Single-Sheet or Multi-Sheet**: By default, it converts only the first sheet to JSON. (Optionally configurable to process all sheets.)  
- **Pretty-Printed JSON**: Generates human-readable JSON with 2-space indentation.  
- **Simple GUI**: Point-and-click interface to choose an Excel file and save the resulting JSON.  
- **Drag-Drop Ready**: Can be hosted in any folder and run by double-clicking the EXE.

---

## Installation

If you want to build from source:

1. **Clone this repository**  
   ```bash
   git clone https://github.com/singhabhudaya/excel-json-converter.git
   cd excel-json-converter
