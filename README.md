# PowerPoint Text Manipulation Script

This script allows you to modify text properties (font size, title font size, and boldness) in PowerPoint files. You can process either a single file or all .pptx files in a directory.

## Features
- Update font size for all text or titles specifically.
- Apply bold formatting to all text.
- Process a single .pptx file or batch process a directory of files.
- Interactive mode for users who don't pass command-line arguments.

## Prerequisites
1. Python: Version 3.7 or higher.
2. Install required libraries using pip:
```bash
pip install -r requirements.txt
```

## Installation
1. Clone the repository or download the script files.
2. Install dependencies by running the provided shell script:
```bash
./install_requirements.sh
```

## Usage
The script can be used either through the command line or interactively.

### Command-Line Usage
Run the script with the following options:

```bash
python script.py -i INPUT -o OUTPUT [OPTIONS]
```

#### Options:
    -i, --input (required): Input file or directory path.
    -o, --output (required): Output file or directory path.
    -d, --directory: Process all .pptx files in the input directory.
    -s, --single: Process a single .pptx file.
    -f, --font_size: Font size for all text.
    -t, --title_font_size: Font size for title text.
    -b, --bold: Apply bold formatting (True or False). 
    
#### Examples

##### Single File
    python script.py -i input.pptx -o output.pptx -s -f 24 -t 32 -b True

##### Directory
    python script.py -i input_folder -o output_folder -d -f 20 -b False

##### Interactive Mode
If no arguments are provided, the script will prompt the user for required inputs interactively:

    Enter the input file or directory path: /path/to/input
    Enter the output file or directory path: /path/to/output
    Is the input a directory? (y/n): y
    Enter the font size for all text (or press Enter to skip): 24
    Enter the font size for title text (or press Enter to skip): 32
    Should all text be bold? (y/n): n

## Requirements
Dependencies are listed in requirements.txt:

python-pptx
Install them using:

    pip install -r requirements.txt

Alternatively, use the provided shell script:

    ./install_requirements.sh

## Testing

To test this the test directories are included in this project and it should produce the `expected_output` folder.

    python3 pptx_editor.py -i testpptx -o testoutput -d -f 48 -t 54 -b false

## Notes
The output directory will be created if it does not exist.
The script identifies the title text as the first line of the first slide with underlined formatting.

## Author
Developed by Ryan Smith.