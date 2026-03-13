# legacy-converter

Convert legacy Microsoft Office files to modern formats using Python and `pywin32`.

## Supports

| Legacy | Modern |
|--------|--------|
| `.doc` | `.docx` |
| `.xls` | `.xlsx` |
| `.ppt` | `.pptx` |
| `.dot` | `.docx` |
| `.xlt` | `.xlsx` |
| `.pot` | `.pptx` |

## Requirements

- Windows with Microsoft Office installed
- Python 3.8+
- `pip install pywin32`

## Usage

```bash
# Convert a whole folder
python legacy_converter.py --input C:\path\to\old_files --output C:\path\to\new_files

# Recurse into subfolders
python legacy_converter.py --input C:\path\to\old_files --output C:\path\to\new_files --recursive

# Single file
python legacy_converter.py --file C:\path\to\file.doc
```

## Notes

- Output files are placed alongside originals if `--output` is not specified
- Already-converted files are skipped automatically
- Logs success/failure for each file
