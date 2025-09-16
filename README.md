# Google Docs Compiler to DOCX

A Python utility to compile multiple Google Docs linked in a DOCX into a single DOCX document, preserving text formatting and inline images.

## Features

1. Extracts hyperlinks from a source DOCX in order
1. Detects and downloads Google Docs links
1. Copies text formatting and inline images
1. Compiles all documents into a single DOCX
1. Supports temporary directories for downloaded files

## Requirements

Python 3.13 or higher



Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Prepare a DOCX with Google Docs hyperlinks.

Run the application:
```bash
python ./main.py -s list-of-urls.docx -o compiled.docx
```

Arguments

`--source, -s`: Path to source DOCX

`--output, -o`: Path for compiled DOCX

`--temp-dir, -t`: Temporary directory (optional). Default is `./temp_dir`

## How It Works

1. Extract hyperlinks using DocxUtils
1. Detect document type per link found (currently only supports Google Docs)
1. Download Google Docs as DOCX
1. Copy text, formatting, and inline images
1. Save the compiled DOCX in order of appearance

## Extending

1. Add support for other document types
1. Extend DocxUtils to handle tables, headers, or footers

## License

### MIT License