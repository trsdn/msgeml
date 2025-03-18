# Email to Text Converter

A Python script to convert .eml and .msg email files to plain text format. The tool also creates a combined chronologically ordered output file.

## Features

- Converts .msg (Outlook) files to plain text
- Converts .eml files to plain text
- Adds date information at the top of each file
- Creates a combined output file with all messages in chronological order
- Preserves original folder structure
- Handles timezone differences for proper chronological sorting

## Requirements

- Python 3.6+
- Required packages:
  - `extract_msg`
  - `beautifulsoup4`
  - `python-dateutil`

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/msgeml.git
cd msgeml
```

2. Install required packages:
```bash
pip install extract_msg beautifulsoup4 python-dateutil
```

## Usage

1. Place your .eml and/or .msg files in the `input` folder
2. Run the script:
```bash
python eml_to_text.py input output
```
3. The converted text files will be created in the `output` folder
4. A combined chronological file named `output.txt` will also be created in the output folder

## File Format

Each converted file includes:
- Date information at the top
- Full message content
- Original formatting where possible

The combined output file includes:
- A separator between messages
- Source filename reference
- Date information
- Full message content in chronological order

## Limitations

- PDF and DOCX attachments are not processed
- Some HTML formatting may be lost in the conversion

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.