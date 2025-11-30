# Word Combiner

A simple utility to combine multiple Word documents (.docx) into a single document with formatting preservation.

Repository: https://github.com/kws/word-combiner

## Features

- Combine multiple .docx files into one document
- Preserves formatting (bold, italic, fonts, colors, alignment)
- Handles tables and complex document structures
- Flexible separation options between documents
- Sort files by name or modification date
- Simple command-line interface

## Installation

### Install with pipx (Recommended)

The easiest way to install `word-combiner` is using [pipx](https://pipx.pypa.io/), which installs Python applications in isolated environments:

```bash
pipx install git+https://github.com/kws/word-combiner.git
```

To upgrade to the latest version:

```bash
pipx upgrade word-combiner
```

### Development Installation

For development, clone the repository and install with Poetry:

```bash
git clone https://github.com/kws/word-combiner.git
cd word-combiner
poetry install
```

## Usage

After installation, the `word-combiner` command will be available:

```bash
word-combiner file1.docx file2.docx file3.docx
```

### Basic Examples

Combine multiple files:
```bash
word-combiner document1.docx document2.docx document3.docx
```

Combine all .docx files in current directory:
```bash
word-combiner *.docx
```

Specify custom output file:
```bash
word-combiner file1.docx file2.docx -o merged.docx
```

### Options

#### `-o, --output PATH`
Specify the output file path. If not provided, defaults to `combined.docx` in the current directory.

```bash
word-combiner file1.docx file2.docx -o output.docx
```

#### `--separator {page_break|newline|none}`
Choose how documents are separated in the combined file:
- `page_break` (default): Insert a page break between documents
- `newline`: Insert a paragraph break between documents
- `none`: No separator between documents

```bash
word-combiner file1.docx file2.docx --separator newline
```

#### `--sort {name|date}`
Sort input files before combining:
- `name`: Sort alphabetically by filename (case-insensitive)
- `date`: Sort by last modified date (oldest first)

```bash
# Sort by name
word-combiner *.docx --sort name

# Sort by modification date
word-combiner *.docx --sort date
```

### Complete Examples

Combine files sorted by name with page breaks:
```bash
word-combiner *.docx --sort name -o combined.docx
```

Combine files sorted by date with newline separators:
```bash
word-combiner file1.docx file2.docx file3.docx --sort date --separator newline -o output.docx
```

## How It Works

The utility:
1. Reads each input .docx file
2. Copies all paragraphs, preserving formatting (bold, italic, fonts, colors, alignment)
3. Copies all tables with their content
4. Inserts separators between documents (if specified)
5. Saves everything to the output file

## Requirements

- Python 3.10 or higher
- Poetry (for dependency management)

## Dependencies

- `click` - Command-line interface framework
- `python-docx` - Library for working with .docx files

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Kaj Wik Siebert (kaj@k-si.com)
