# Metadata Extractor

Extract metadata from various file formats including PDFs, images (PNG, JPEG, TIFF, BMP, GIF), and Office documents (DOCX, XLSX, PPTX).

## Installation

Before running the script, you need to install the required libraries. Use the following command:

```bash
pip install PyPDF2 Pillow python-docx openpyxl python-pptx hachoir
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

* Detect the format:
```bash
python Metalookup.py -f /path/to/single/file.pdf -D 
```

* Help:
```bash
❯ python3 Metalookup.py -h

  __  __     _          _            _             
 |  \/  |___| |_ __ _  | |   ___  __| |___  _ _ __ 
 | |\/| / -_)  _/ _` | | |__/ _ \/ _| / / || | '_ \
 |_|  |_\___|\__\__,_| |____\___/\__|_\_\\_,_| .__/
                                             |_|   

usage: Metalookup.py [-h] [-v] (-f FILE | -d DIRECTORY) [-D]

Extract metadata from files.

options:
  -h, --help            show this help message and exit
  -v, --version         show program's version number and exit
  -f FILE, --file FILE  File to extract metadata from or detect its type
  -d DIRECTORY, --directory DIRECTORY
                        Directory to extract metadata from all contained files
  -D, --detect          Detect the file type based on its magic number. Requires -f.
```

## Features

* PDF Metadata Extraction: Extracts information from the properties of a PDF.
* Image Metadata Extraction: Grabs general information as well as EXIF data (commonly used by cameras).
* Office Documents Metadata Extraction: Extracts details from Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) files.

## Contributing
If you'd like to contribute, please fork the repository and use a feature branch. Pull requests are warmly welcome.

## Licensing

This project is licensed under MIT license. 

