# word2pdf

A CLI toolkit for automating Word document processing tasks: converting Word files to PDF, extracting template tags, and batch processing entire folders.

## Background

This project was built to automate a repetitive manual workflow that previously required dedicated developer time for each occurrence. What used to be a recurring task assigned as a story point is now a simple command-line operation, freeing up valuable developer time for the entire team.

## Features

- **Word to PDF Conversion**: Convert single files or entire folders of Word documents to PDF
- **Template Tag Extraction**: Extract `«placeholder»` tags from mail-merge templates and export them to a text file or CSV
- **Batch Processing**: Process multiple documents at once with progress tracking
- **Clipboard Integration**: Automatically copies output file paths to clipboard for quick access

## Requirements

- **Windows OS** with Microsoft Word installed
- Python 3.x

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/word2pdf.git
   cd word2pdf
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Convert Word to PDF

Convert a single file:
```bash
python word2pdf.py "path/to/document.docx"
```

Convert a single file to a specific output location:
```bash
python word2pdf.py "path/to/document.docx" "path/to/output.pdf"
```

Bulk convert all Word files in a folder:
```bash
python word2pdf.py "path/to/folder" "path/to/output/folder"
```

### Extract Template Tags

Extract all `«tagname»` placeholders from Word templates:

```bash
python getTags.py "path/to/template.docx"
```

Process an entire folder of templates:
```bash
python getTags.py "path/to/folder"
```

This generates an `Info.txt` file listing all unique tags found in each document.

### Replace Newlines with HTML Tags

Append `<br>` tags to paragraph endings (useful for HTML export preparation):

```bash
python replaceNewLine.py "path/to/document.docx"
```

### Generate Tag Matrix (Jupyter Notebook)

Use `createCSV.ipynb` to analyze a folder of templates and generate a CSV matrix showing which tags appear in which documents. Useful for template auditing and documentation.

## Project Structure

```
word2pdf/
├── word2pdf.py        # Main Word-to-PDF converter
├── getTags.py         # Template tag extractor
├── replaceNewLine.py  # Newline to <br> converter
├── helpers.py         # Shared utility functions
├── createCSV.ipynb    # Jupyter notebook for tag analysis
└── requirements.txt   # Python dependencies
```

## How It Works

The toolkit uses Python's `win32com.client` to interface with Microsoft Word's COM automation API. This allows programmatic control over Word to:

- Open documents in various formats (`.doc`, `.docx`)
- Save documents in different formats (PDF, DOCX)
- Read document content for tag extraction

The `python-docx` library is used for direct `.docx` file manipulation when COM automation isn't required.
