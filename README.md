# DocDown

DocDown is a command-line tool that converts Microsoft Word documents (.doc/.docx) to Markdown (.md) format. It handles both single files and directories, preserving document structure and images.

## Features

- Converts Word documents (.docx) to Markdown format
- Extracts and saves images to a separate folder
- Supports heading styles
- Handles both single files and directories
- Maintains directory hierarchy when processing folders
- Comprehensive logging for better debugging
- Cross-platform compatibility

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/docdown.git
   cd docdown
   ```

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows, use: venv\Scripts\activate
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Converting a Single File

```bash
python docdown.py input.docx output_directory/
```

### Converting All Files in a Directory

```bash
python docdown.py input_directory/ output_directory/
```

### Specifying Custom Log Directory

```bash
python docdown.py input_directory/ output_directory/ --log-dir custom_logs
```

## Output Structure

The tool creates the following structure in your output directory, maintaining the same hierarchy as the source:

```
output_directory/
├── subfolder1/
│   ├── images/
│   │   ├── document1_image_1.png
│   │   └── document1_image_2.jpg
│   └── document1.md
├── subfolder2/
│   ├── images/
│   │   └── document2_image_1.png
│   └── document2.md
└── logs/
    └── docdown_20240321_143022.log
```

- Markdown files are placed in the same relative directory structure as the source
- Images are extracted and saved in an `images/` subdirectory within each folder
- Logs are saved in the specified log directory (default: `logs/`)
- Image references in the Markdown files point to the correct relative image locations

## Command Line Arguments

- `source`: Path to the input file or directory (required)
- `target`: Path to the output directory (required)
- `--log-dir`: Directory for log files (optional, default: 'logs')

## Current Limitations

- Only supports .docx format (newer Word format)
- Basic formatting support (headings and paragraphs)
- Tables and complex formatting are not yet supported

## Future Improvements

- Support for tables
- Support for lists (ordered and unordered)
- Text formatting (bold, italic, etc.)
- Hyperlinks support
- Document metadata
- Page breaks
- Footnotes
- Support for older .doc format

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

