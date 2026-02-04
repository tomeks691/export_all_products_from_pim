# Strapi Exporter GUI

A simple desktop GUI to export all products from Strapi into a single XLSX file.

## Features
- Exports all products with pagination
- Supports Strapi v4 and v5 API shapes
- Flattens `attributes` from Strapi v4 (optional)
- Splits `rozmiary` into separate rows
- Appends size to `sku` (`SKU/ROZMIAR`)
- Saves settings locally to `settings.json`

## Requirements
- Python 3.10+

## Installation
```bash
python -m venv .venv
# Windows
.\.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Usage
```bash
python strapi_export_gui.py
```

## Settings
The app stores the last used settings in `settings.json` in the project directory.

## Build EXE (GitHub Actions)
A GitHub Actions workflow is included to build a Windows EXE on release and attach it to the release assets.

## Notes
If you use Strapi v4, keep the option "Flatten attributes" enabled.
