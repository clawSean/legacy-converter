"""
legacy_converter.py
Converts legacy Microsoft Office files to modern formats:
  .doc  → .docx
  .xls  → .xlsx
  .ppt  → .pptx

Requirements:
  - Windows with Microsoft Office installed
  - pip install pywin32

Usage:
  python legacy_converter.py --input C:\path\to\files --output C:\path\to\output
  python legacy_converter.py --input C:\path\to\files  # outputs alongside originals
  python legacy_converter.py --file C:\path\to\file.doc  # single file
"""

import os
import sys
import argparse
import logging
from pathlib import Path

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger(__name__)

# Format mappings: extension → (Office app, save format constant)
FORMAT_MAP = {
    ".doc":  ("word",       16),   # wdFormatXMLDocument = 16 (.docx)
    ".xls":  ("excel",      51),   # xlOpenXMLWorkbook = 51 (.xlsx)
    ".ppt":  ("powerpoint", 24),   # ppSaveAsOpenXMLPresentation = 24 (.pptx)
    ".dot":  ("word",       16),   # Word template → .docx
    ".xlt":  ("excel",      51),   # Excel template → .xlsx
    ".pot":  ("powerpoint", 24),   # PowerPoint template → .pptx
}

NEW_EXT = {
    ".doc": ".docx",
    ".xls": ".xlsx",
    ".ppt": ".pptx",
    ".dot": ".docx",
    ".xlt": ".xlsx",
    ".pot": ".pptx",
}

# Trust Center File Block registry paths (Office 16.0 / 365)
_TRUST_CENTER_PATHS = {
    "word":       r"Software\Microsoft\Office\16.0\Word\Security\FileBlock",
    "excel":      r"Software\Microsoft\Office\16.0\Excel\Security\FileBlock",
    "powerpoint": r"Software\Microsoft\Office\16.0\PowerPoint\Security\FileBlock",
}

# Maps file extension → (app, registry value name)
# These control which legacy formats are blocked in the Trust Center
_FILE_BLOCK_VALUES = {
    ".doc": ("word",       "BinaryFiles"),
    ".dot": ("word",       "BinaryFiles"),
    ".xls": ("excel",      "XL9597WorkbooksAndTemplates"),
    ".xlt": ("excel",      "XL9597WorkbooksAndTemplates"),
    ".ppt": ("powerpoint", "PowerPoint972003Presentations"),
    ".pot": ("powerpoint", "PowerPoint972003Presentations"),
}


def _set_trust_center_file_block(ext: str, block: bool) -> tuple:
    """
    Temporarily allow or restore a Trust Center File Block setting.

    Returns (key_path, value_name, original_value) on success,
    or (None, None, None) if registry access failed or key not applicable.

    block=False → set to 0 (allow open)
    block=True  → restore original value
    """
    try:
        import winreg
    except ImportError:
        return None, None, None

    if ext not in _FILE_BLOCK_VALUES:
        return None, None, None

    app_name, value_name = _FILE_BLOCK_VALUES[ext]
    key_path = _TRUST_CENTER_PATHS[app_name]

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
    except OSError:
        # Key doesn't exist yet — no block is set, nothing to do
        return None, None, None

    try:
        original, _ = winreg.QueryValueEx(key, value_name)
    except OSError:
        original = None  # Value not present → no block configured

    try:
        if not block:
            winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, 0)
        winreg.CloseKey(key)
        return key_path, value_name, original
    except OSError as e:
        winreg.CloseKey(key)
        log.warning(f"Could not modify Trust Center File Block (may be GPO-enforced): {e}")
        return None, None, None


def _restore_trust_center_file_block(key_path: str, value_name: str, original):
    """Restore a previously modified Trust Center File Block registry value."""
    if key_path is None:
        return
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        if original is None:
            try:
                winreg.DeleteValue(key, value_name)
            except OSError:
                pass
        else:
            winreg.SetValueEx(key, value_name, 0, winreg.REG_DWORD, original)
        winreg.CloseKey(key)
    except Exception:
        pass


def convert_file(src: Path, dest_dir: Path = None) -> Path | None:
    """Convert a single legacy Office file to its modern equivalent."""
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        log.error("pywin32 not installed. Run: pip install pywin32")
        sys.exit(1)

    ext = src.suffix.lower()
    if ext not in FORMAT_MAP:
        log.warning(f"Skipping unsupported format: {src.name}")
        return None

    app_name, save_format = FORMAT_MAP[ext]
    new_ext = NEW_EXT[ext]

    if dest_dir:
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / (src.stem + new_ext)
    else:
        dest = src.with_suffix(new_ext)

    # Skip if output already exists
    if dest.exists():
        log.info(f"Already exists, skipping: {dest.name}")
        return dest

    # Temporarily disable Trust Center File Block for this extension
    key_path, value_name, original = _set_trust_center_file_block(ext, block=False)

    pythoncom.CoInitialize()
    app = None
    try:
        log.info(f"Converting: {src.name} → {dest.name}")

        if app_name == "word":
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            doc = app.Documents.Open(str(src.resolve()))
            doc.SaveAs2(str(dest.resolve()), FileFormat=save_format)
            doc.Close()

        elif app_name == "excel":
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            wb = app.Workbooks.Open(str(src.resolve()))
            wb.SaveAs(str(dest.resolve()), FileFormat=save_format)
            wb.Close()

        elif app_name == "powerpoint":
            app = win32com.client.Dispatch("PowerPoint.Application")
            prs = app.Presentations.Open(str(src.resolve()), WithWindow=False)
            prs.SaveAs(str(dest.resolve()), FileFormat=save_format)
            prs.Close()

        log.info(f"  ✓ Saved: {dest.name}")
        return dest

    except Exception as e:
        log.error(f"  ✗ Failed: {src.name} — {e}")
        return None

    finally:
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        # Always restore the original File Block setting
        _restore_trust_center_file_block(key_path, value_name, original)


def convert_directory(input_dir: Path, output_dir: Path = None, recursive: bool = False):
    """Convert all legacy Office files in a directory."""
    pattern = "**/*" if recursive else "*"
    files = [f for f in input_dir.glob(pattern) if f.suffix.lower() in FORMAT_MAP and f.is_file()]

    if not files:
        log.info("No legacy Office files found.")
        return

    log.info(f"Found {len(files)} file(s) to convert.")
    success, failed = 0, 0

    for f in files:
        # Mirror directory structure if recursive + output_dir set
        if output_dir and recursive:
            rel = f.parent.relative_to(input_dir)
            dest = output_dir / rel
        else:
            dest = output_dir

        result = convert_file(f, dest)
        if result:
            success += 1
        else:
            failed += 1

    log.info(f"\nDone. ✓ {success} converted, ✗ {failed} failed.")


def main():
    parser = argparse.ArgumentParser(description="Convert legacy Office files to modern formats.")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--input", "-i", type=Path, help="Input directory")
    group.add_argument("--file", "-f", type=Path, help="Single input file")
    parser.add_argument("--output", "-o", type=Path, default=None,
                        help="Output directory (default: same as input)")
    parser.add_argument("--recursive", "-r", action="store_true",
                        help="Recurse into subdirectories")
    args = parser.parse_args()

    if args.file:
        if not args.file.exists():
            log.error(f"File not found: {args.file}")
            sys.exit(1)
        convert_file(args.file, args.output)

    elif args.input:
        if not args.input.is_dir():
            log.error(f"Directory not found: {args.input}")
            sys.exit(1)
        convert_directory(args.input, args.output, args.recursive)


if __name__ == "__main__":
    main()
