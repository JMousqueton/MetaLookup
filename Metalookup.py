import argparse
import os
from PyPDF2 import PdfReader
from PIL import Image, UnidentifiedImageError, ExifTags
from docx import Document
from openpyxl import load_workbook
from pptx import Presentation
from hachoir.metadata import extractMetadata
from hachoir.parser import createParser

Version = "0.3.0"

Banner = r"""
  __  __     _          _            _             
 |  \/  |___| |_ __ _  | |   ___  __| |___  _ _ __ 
 | |\/| / -_)  _/ _` | | |__/ _ \/ _| / / || | '_ \
 |_|  |_\___|\__\__,_| |____\___/\__|_\_\\_,_| .__/
                                             |_|   
"""

MAGIC_NUMBERS = {
    # Images
    b"\x89PNG\r\n\x1a\n": "PNG",
    b"\xff\xd8": "JPEG",
    b"\x47\x49\x46\x38\x37\x61": "GIF87a",
    b"\x47\x49\x46\x38\x39\x61": "GIF89a",
    b"\x42\x4D": "BMP",
    b"\x00\x00\x01\x00": "ICO",
    b"\x00\x00\x02\x00": "CUR",
    b"\x49\x49\x2A\x00": "TIFF (little-endian)",
    b"\x4D\x4D\x00\x2A": "TIFF (big-endian)",
    b"\x38\x42\x50\x53": "PSD",
    b"\x77\x45\x42\x50": "WebP",

    # Videos
    b"\x00\x00\x00\x1C\x66\x74\x79\x70": "MP4",
    b"\x1A\x45\xDF\xA3": "MKV or WebM",
    b"\x52\x49\x46\x46": "AVI",
    b"\x00\x00\x01\xBA": "MPEG",
    b"\x00\x00\x01\xB3": "MPEG",
    b"\x66\x74\x79\x70\x33\x67\x70": "3GP",
    b"\x4F\x67\x67\x53": "OGG/OGV",
    b"\x46\x4C\x56\x01": "FLV",
    b"\x52\x49\x46\x46": "RIFF (could be AVI or WAV)",

    # Documents (you have these already, but I'm adding just for context)
    b"\x25\x50\x44\x46": "PDF",
    b"\x50\x4B\x03\x04": "ZIP or Office Document (.docx, .xlsx, etc.)"
}


def detect_file_type(file_path):
    max_len = max(len(magic) for magic in MAGIC_NUMBERS.keys())
    with open(file_path, 'rb') as file:
        file_start = file.read(max_len)
        for magic, filetype in MAGIC_NUMBERS.items():
            if file_start.startswith(magic):
                return filetype
    return "Unknown"

def safe_getattr(obj, attr, default=None):
    return getattr(obj, attr, default)

def extract_pdf_metadata(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf = PdfReader(file)
        info = pdf.trailer["/Info"]
        metadata = {key[1:]: info[key] for key in info if key != '/ID'}
        return metadata

def extract_video_metadata(video_path):
    parser = createParser(video_path)
    if not parser:
        print(f"Unable to parse video file: {video_path}")
        return {}
    
    with parser:
        try:
            metadata = extractMetadata(parser)
        except Exception as e:
            print(f"Metadata extraction error: {e}")
            return {}
    
    if not metadata:
        return {}

    # Iterate over metadata attributes and fetch their values
    meta_data_dict = {}
    for item in metadata:
        if item.values:
            meta_data_dict[item.key] = item.values[0].value

    return meta_data_dict


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
    elif file_path.lower().endswith(('.mp4', '.mkv')):
        return extract_video_metadata(file_path)
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
    print(Banner)
    parser = argparse.ArgumentParser(description='Extract metadata from files.')
    parser.add_argument('-v', '--version', action='version', version=f"Metalookup Version {Version}")
    
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('-f', '--file', help='File to extract metadata from or detect its type', type=str)
    group.add_argument('-d', '--directory', help='Directory to extract metadata from all contained files', type=str)

    # This is the new option for detection
    parser.add_argument('-D', '--detect', action='store_true', help="Detect the file type based on its magic number. Requires -f.")

    args = parser.parse_args()

    if args.file:
        if not os.path.exists(args.file):
            print(f"Error: The file '{args.file}' does not exist.")
        elif args.detect:  # New detection logic
            file_type = detect_file_type(args.file)
            print(f"The file {args.file} appears to be of type: {file_type}")
        else:
            metadata = extract_metadata(args.file)
            print(f"Metadata for {args.file}:\n{metadata}\n")
    elif args.directory:
        extract_metadata_from_directory(args.directory)

