# AL Bank Statement Extractor

Converts Arbejdernes Landsbank PDF statements to Excel (.xlsx), automatically handling Danish number formatting and column parsing.

## Installation

    pip install pdfplumber pandas openpyxl

## Usage

### Basic Extraction
    python al-bank-extractor-final.py -f Statement.pdf

### Extraction with Cleanup
Removes redundant dates from the description (e.g., "Text 05.05" -> "Text").

    python al-bank-extractor-final.py -f Statement.pdf --clean

## Troubleshooting

If python isn't recognized or modules appear missing, use the Windows py launcher:

    py -m pip install pdfplumber pandas openpyxl
    py al-bank-extractor-final.py -f Statement.pdf
