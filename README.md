# Document Translation Utility

A Python module for translating Excel, Word, PDF, CSV, and plain text documents from one language to another. Features intelligent caching to avoid redundant API calls and preserve translation consistency across sessions.

## Features

- **Multi-format support**: Translate Excel (.xlsx, .xls), Word (.docx), PDF (.pdf), CSV (.csv), and plain text (.txt) files
- **Batch processing**: Translate entire directories with optional recursion
- **Translation cache**: Persistent cache system to avoid redundant API calls and ensure consistency
- **Unique value optimization**: Extracts and translates only unique text values, then maps back to reduce redundant processing
- **Directory structure preservation**: Automatically recreates source directory structure in target location
- **Smart filename translation**: Translates both directory names and filenames
- **Resume capability**: Skip already-translated files and reuse cached translations
- **Comprehensive logging**: Track progress and errors during translation

## Installation

### Required Dependencies

```bash
pip install pandas openpyxl deep-translator
```

### Optional Dependencies

For Word document support:
```bash
pip install python-docx
```

For PDF document support:
```bash
pip install pypdf
```

Or install all dependencies at once:
```bash
pip install pandas openpyxl deep-translator python-docx pypdf
```

## Quick Start

```python
from translate_files import translate_directory
import logging

# Set up logging to see progress
logging.basicConfig(level=logging.INFO)

# Translate a directory
translate_directory(
    source_dir="/path/to/source/directory",
    target_dir="/path/to/target/directory",
    source_lang='th',      # Thai
    target_lang='en',      # English
    recursive=True         # Process subdirectories
)
```

## Supported File Formats

### Excel Files (.xlsx, .xls)

- Treats first row as data (no column headers)
- Translates all cell values including the first row
- **Workbook-level optimization**: Extracts unique values across ALL sheets, translates once
- Preserves data structure and formatting
- Handles null/NaN values gracefully
- Translates sheet names

### CSV Files (.csv)

- Translates column headers separately
- Translates all cell values using unique value extraction
- Automatically detects CSV delimiter
- Handles various text encodings (UTF-8, latin-1)
- Preserves data structure

### Plain Text Files (.txt)

- Translates entire text content
- Preserves line breaks and paragraph structure
- Handles various text encodings automatically
- Output saved as UTF-8

### Word Documents (.docx)

- Translates paragraphs
- Translates table cells
- Preserves document structure, formatting, and layout
- Requires `python-docx` package

### PDF Documents (.pdf)

- Extracts and translates text from pages
- **Note**: PDF text replacement has limitations due to the complex nature of PDF format
- Structure preservation may vary
- Requires `pypdf` or `PyPDF2` package
```

## Usage Examples

### Basic Usage

Translate all supported files in a directory:

```python
from translate_files import translate_directory

translate_directory(
    source_dir="data/thai_documents",
    target_dir="data/english_documents",
    source_lang='th',
    target_lang='en',
    recursive=True
)
```

### Excel Files Only

Translate only Excel files, non-recursive:

```python
translate_directory(
    source_dir="data/excel_files",
    target_dir="data/translated",
    source_lang='th',
    target_lang='en',
    recursive=False,
    file_extensions=('.xlsx', '.xls')
)
```

### Custom Cache Location

Specify a custom cache file location:

```python
translate_directory(
    source_dir="data/source",
    target_dir="data/target",
    source_lang='th',
    target_lang='en',
    cache_file="custom_cache.json",
    recursive=True
)
```

### Translate Single File

```python
from translate_files import translate_file, TranslationCache
from deep_translator import GoogleTranslator

cache = TranslationCache('cache.json')
translator = GoogleTranslator(source='th', target='en')

translate_file(
    input_path="document.xlsx",
    output_path="document_translated.xlsx",
    translator=translator,
    cache=cache
)
```

## API Reference

### `translate_directory()`

Main function for batch translation of files in a directory.

**Parameters:**

- `source_dir` (str): Path to the source directory containing files to translate
- `target_dir` (str): Path to the target directory where translated files will be saved
- `source_lang` (str, optional): Source language code. Default: `'th'` (Thai)
- `target_lang` (str, optional): Target language code. Default: `'en'` (English)
- `cache_file` (str, optional): Path to the translation cache file. If `None`, defaults to `translation_cache.json` in the target directory
- `recursive` (bool, optional): If `True`, recursively processes subdirectories. Default: `True`
- `file_extensions` (tuple, optional): Tuple of file extensions to process (case-insensitive). Default: `('.xlsx', '.xls', '.docx', '.pdf', '.csv', '.txt')`

**Example:**

```python
translate_directory(
    source_dir="/data/thai_docs",
    target_dir="/data/english_docs",
    source_lang='th',
    target_lang='en',
    recursive=True,
    file_extensions=('.xlsx', '.docx', '.csv', '.txt')
)
```

### `translate_file()`

Translate a single file based on its extension.

**Parameters:**

- `input_path` (str): Path to the source file
- `output_path` (str): Path where the translated file will be saved
- `translator` (GoogleTranslator): Configured translator instance
- `cache` (TranslationCache): Translation cache instance

**Supported formats:** `.xlsx`, `.xls`, `.docx`, `.pdf`, `.csv`, `.txt`

### Format-Specific Functions

- `translate_excel()`: Translate Excel files (all cells including first row, no headers, uses unique value extraction)
- `translate_csv()`: Translate CSV files (column headers + cell values using unique value extraction)
- `translate_txt()`: Translate plain text files (preserves line breaks)
- `translate_word()`: Translate Word documents (paragraphs and table cells)
- `translate_pdf()`: Translate PDF documents (with limitations - see Notes)
- `translate_dataframe_values()`: Helper function for efficient DataFrame translation using unique value extraction

### `TranslationCache`

Class for managing persistent translation cache.

**Methods:**

- `get(text: str) -> Optional[str]`: Retrieve cached translation
- `set(text: str, translation: str) -> None`: Store translation in cache
- `save() -> None`: Save cache to disk

## Translation Cache System

The translation cache is a persistent JSON file that stores previously translated text strings. This provides several benefits:

### Benefits

1. **Performance**: Avoids redundant API calls for text that has already been translated, significantly speeding up batch translation operations.

2. **Cost Efficiency**: Reduces API usage costs by reusing existing translations.

3. **Consistency**: Ensures the same source text always translates to the same target text across different translation sessions.

4. **Resume Capability**: If a translation process is interrupted, previously translated text can be reused without re-translating.

### How It Works

- The cache is automatically loaded when a `TranslationCache` instance is created
- New translations are stored in memory and saved to disk every 100 translations
- The cache is saved at the end of each translation session
- Cache file is stored as JSON with UTF-8 encoding to preserve special characters
- The same cache can be shared across multiple translation sessions

### Cache File Location

By default, the cache file is stored as `translation_cache.json` in the target directory. You can specify a custom location using the `cache_file` parameter.

## Unique Value Optimization

For Excel and CSV files, the module uses an efficient translation strategy:

1. **Extract unique values**: Collects all unique text values across the entire file
2. **Translate once**: Each unique value is translated only once (leveraging the cache)
3. **Map back**: Translations are mapped back to original positions using fast vectorized operations

### Excel Workbook-Level Optimization

For Excel files, unique values are extracted across **all sheets** in the workbook, not just per-sheet. This means:
- Values shared between sheets are translated only once
- Sheet names are included in the same deduplication pass
- A single translation dictionary is built for the entire workbook

### Benefits

- **Reduced function calls**: For a column with 10,000 rows but only 10 unique values, only 10 translation calls are made instead of 10,000
- **Lower overhead**: Avoids per-cell type checking and cache lookups for duplicate values
- **Cross-column deduplication**: Duplicate values across different columns are also handled efficiently
- **Cross-sheet deduplication** (Excel): Values appearing in multiple sheets are translated only once

This optimization is particularly effective for datasets with categorical data, repeated labels, or standardized values.


## Language Support

The module uses Google Translate via the `deep-translator` library, which supports a wide range of languages. Common language codes:

- `'th'`: Thai
- `'en'`: English
- `'zh'`: Chinese
- `'ja'`: Japanese
- `'ko'`: Korean
- `'es'`: Spanish
- `'fr'`: French
- `'de'`: German

See the [deep-translator documentation](https://github.com/nidhaloff/deep-translator) for the full list of supported languages.

## Logging

The module uses Python's `logging` module to track progress and errors. Set up logging to see what's happening:

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
```

## Notes and Limitations

1. **PDF Translation**: PDF text replacement is limited. The function extracts text and attempts to preserve structure, but full text replacement in PDFs is complex. Consider using specialized PDF editing tools for production use.

2. **Rate Limiting**: The module includes a 0.5-second delay between API calls to avoid rate limiting. For large batches, this may result in longer processing times.

3. **Thai Character Detection**: The module only translates text containing Thai characters (Unicode range U+0E00-U+0E7F). Other text is returned as-is to avoid unnecessary API calls.

4. **File Skipping**: Files that already exist in the target location are automatically skipped to allow resuming interrupted translation jobs.

5. **Directory Names**: Directory names in the source path are also translated when recreating the directory structure in the target location.

## Error Handling

The module includes comprehensive error handling:

- Missing dependencies raise `ImportError` with installation instructions
- Unsupported file formats raise `ValueError`
- File read/write errors are logged and processing continues with other files
- Translation API errors are logged and the original text is returned

## Requirements

- Python 3.7+
- pandas
- openpyxl
- deep-translator
- python-docx (optional, for Word files)
- pypdf or PyPDF2 (optional, for PDF files)

## License

This module is provided as-is for use in document translation workflows.

## Contributing

Contributions are welcome! Please ensure that:

- Code follows the existing style and patterns
- All functions are properly documented
- Error handling is comprehensive
- Type hints are included where appropriate

