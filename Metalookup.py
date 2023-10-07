import argparse
import os
from PyPDF2 import PdfReader
from PIL import Image, UnidentifiedImageError, ExifTags
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation


def safe_getattr(obj, attr, default=None):
    return getattr(obj, attr, default)

def extract_pdf_metadata(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf = PdfReader(file)
        info = pdf.trailer["/Info"]
        metadata = {key[1:]: info[key] for key in info if key != '/ID'}
        return metadata

def extract_image_metadata(img_path):
    try:
        with Image.open(img_path) as img:
            info = img.info
            
            # Check if the image has the "exif" attribute
            if "exif" in info and info["exif"]:
                exif_data = img._getexif()
                
                if exif_data:
                    # Convert the tag ID to the tag name for better readability
                    decoded_exif = {ExifTags.TAGS[k]: v for k, v in exif_data.items() if k in ExifTags.TAGS}
                    info["decoded_exif"] = decoded_exif

            # Remove raw EXIF data from the metadata
            info.pop("exif", None)
            
            metadata = {key: info[key] for key in info if not key.startswith('icc_')}
            return metadata
            
    except UnidentifiedImageError:
        print(f"Error: Could not identify image: {img_path}")
        return {}

def extract_office_metadata(doc_path):
    if doc_path.endswith('.docx'):
        doc = Document(doc_path)
        core_props = doc.core_properties
    elif doc_path.endswith('.xlsx'):
        wb = load_workbook(doc_path)
        core_props = wb.properties
    elif doc_path.endswith('.pptx'):
        pres = Presentation(doc_path)
        core_props = pres.core_properties
    else:
        return {}

    metadata = {
        'author': safe_getattr(core_props, 'author'),
        'category': safe_getattr(core_props, 'category'),
        'comments': safe_getattr(core_props, 'comments'),
        'content_status': safe_getattr(core_props, 'content_status'),
        'created': safe_getattr(core_props, 'created'),
        'identifier': safe_getattr(core_props, 'identifier'),
        'keywords': safe_getattr(core_props, 'keywords'),
        'language': safe_getattr(core_props, 'language'),
        'last_modified_by': safe_getattr(core_props, 'last_modified_by'),
        'last_printed': safe_getattr(core_props, 'last_printed'),
        'modified': safe_getattr(core_props, 'modified'),
        'revision': safe_getattr(core_props, 'revision'),
        'subject': safe_getattr(core_props, 'subject'),
        'title': safe_getattr(core_props, 'title'),
        'version': safe_getattr(core_props, 'version')
    }
    return metadata

def extract_metadata(file_path):
    if file_path.endswith('.pdf'):
        return extract_pdf_metadata(file_path)
    elif file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.gif')):
        return extract_image_metadata(file_path)
    elif file_path.lower().endswith(('.docx', '.xlsx', '.pptx')):
        return extract_office_metadata(file_path)
    else:
        print(f"Unsupported file format for {file_path}")
        return {}

def extract_metadata_from_directory(directory):
    for root, _, files in os.walk(directory):
        for name in files:
            if name.startswith('~$'):
                continue  # Skip files starting with ~$
            file_path = os.path.join(root, name)
            metadata = extract_metadata(file_path)
            print(f"Metadata for {file_path}:\n{metadata}\n")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Extract metadata from files.')
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--file', help='File to extract metadata from', type=str)
    group.add_argument('-d', '--directory', help='Directory to extract metadata from all contained files', type=str)

    args = parser.parse_args()

    if args.file:
        metadata = extract_metadata(args.file)
        print(f"Metadata for {args.file}:\n{metadata}\n")
    elif args.directory:
        extract_metadata_from_directory(args.directory)

