import os
import re
import sys
import subprocess
from pathlib import Path

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package, "--quiet"])

# Ensure pypdf is installed
install_and_import("pypdf")
from pypdf import PdfWriter

def natural_sort_key(s):
    """Sort strings containing numbers in a way that humans expect (page1, page2, page10)."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', str(s))]

def merge_pdfs(search_dir, output_filename):
    search_path = Path(search_dir)
    if not search_path.exists():
        print(f"Error: Directory not found: {search_dir}")
        return

    # Find all PDF files and sort them naturally
    pdf_files = sorted(list(search_path.glob("*.pdf")), key=natural_sort_key)
    
    # Filter out the output file if it already exists to avoid recursive merging
    pdf_files = [f for f in pdf_files if f.name != output_filename]

    if not pdf_files:
        print(f"No PDF files found in: {search_dir}")
        return

    print(f"Found {len(pdf_files)} PDF files. Merging...")
    writer = PdfWriter()
    
    for pdf in pdf_files:
        try:
            print(f"  Adding: {pdf.name}")
            writer.append(str(pdf))
        except Exception as e:
            print(f"  [FAILED] Could not add {pdf.name}: {e}")

    output_path = search_path / output_filename
    
    try:
        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"\n[SUCCESS] All files merged into: {output_path}")
    except Exception as e:
        print(f"\n[ERROR] Failed to save merged PDF: {e}")

if __name__ == "__main__":
    # Automatic path detection
    SCRIPT_PATH = Path(__file__).resolve()
    BASE_DIR = SCRIPT_PATH.parents[2]

    # Default directory
    default_dir = BASE_DIR / "99. Generate Report" / "OUTPUT FILE" / "PDF_REPORTS"

    # Check for argument
    if len(sys.argv) > 1:
        target_dir = Path(sys.argv[1])
        if not target_dir.is_absolute():
            # If not absolute, assume it's a subfolder in OUTPUT FILE
            target_dir = BASE_DIR / "99. Generate Report" / "OUTPUT FILE" / sys.argv[1]
    else:
        target_dir = default_dir

    output_file = "ALL_MERGED_REPORTS.pdf"

    print(f"Target Directory: {target_dir}")
    merge_pdfs(target_dir, output_file)
    input("\nPress Enter to close...")

