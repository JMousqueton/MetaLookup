# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.0] - 2023-10-07

### Added
- Support for extracting metadata from video files (.mp4 and .mkv) using the `hachoir` library.
- Added `-v` or `--version` argument to display the current version of the script.
- Introduced `-D` or `--detect` option as a sub-option for `-f` to determine the file format based on its magic number.
  - Allows for basic file format detection for common types such as PNG, JPEG, PDF, and some Office documents.
  - The detection is performed using the initial bytes (magic numbers) of the file to recognize its type.

### Changed
- Modified the behavior when using the `-f` flag. The script now verifies if the specified file exists before attempting to extract metadata.

## [0.2.0] - 2023-07-21

### Added
- Initial release.
- Support for extracting metadata from PDF files using `PyPDF2`.
- Support for extracting metadata from image formats like PNG, JPG, JPEG, TIFF, BMP, GIF using `PIL`.
- Support for extracting metadata from Office files (.docx, .xlsx, .pptx) using `docx`, `openpyxl`, and `pptx`.