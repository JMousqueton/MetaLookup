# Metadata Extractor

Extract metadata from various file formats including PDFs, images (PNG, JPEG, TIFF, BMP, GIF), and Office documents (DOCX, XLSX, PPTX).

## Installation

Before running the script, you need to install the required libraries. Use the following command:

```bash
pip install argparse PyPDF2 Pillow python-docx openpyxl python-pptx
```
or 
```bash
pip install -r requirements.txt
```

## Usage

The script can extract metadata from a single file or from all files in a directory.

* Extract metadata from a single file:

```bash
python Metalookup.py -f /path/to/single/file.pdf
```

* Extract metadata from all files in a directory:


```bash
python Metalookup.py -d /path/to/directory/
```

## Features

* PDF Metadata Extraction: Extracts information from the properties of a PDF.
* Image Metadata Extraction: Grabs general information as well as EXIF data (commonly used by cameras).
* Office Documents Metadata Extraction: Extracts details from Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) files.

## Contributing
If you'd like to contribute, please fork the repository and use a feature branch. Pull requests are warmly welcome.

## Licensing

This project is licensed under MIT license. 

