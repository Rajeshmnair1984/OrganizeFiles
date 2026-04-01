#!/usr/bin/env python3
import os
import shutil
import argparse
import logging
import sys
import subprocess
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

def install(package):
    """Install a package using pip."""
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to install {package}: {e}")
        return False

# Auto-install dependencies
try:
    from pypdf import PdfReader
    HAS_PDF = True
except ImportError:
    print("Installing pypdf for PDF support...")
    if install("pypdf"):
        try:
            from pypdf import PdfReader
            HAS_PDF = True
        except ImportError:
            HAS_PDF = False
    else:
        HAS_PDF = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    print("Installing python-docx for DOCX support...")
    if install("python-docx"):
        try:
            from docx import Document
            HAS_DOCX = True
        except ImportError:
            HAS_DOCX = False
    else:
        HAS_DOCX = False

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Configuration
FOLDERS_TO_ORGANIZE = [
    Path.home() / "Downloads",
    Path.home() / "Documents"
]

CATEGORIES = {
    "Documents": [".pdf", ".docx", ".doc", ".txt", ".xlsx", ".pptx", ".csv"],
    "Images": [".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".heic"],
    "Videos": [".mp4", ".mov", ".avi", ".mkv"],
    "Archives": [".zip", ".tar", ".gz", ".rar", ".7z"],
    "Music": [".mp3", ".wav", ".flac", ".m4a"],
    "Applications": [".dmg", ".pkg", ".app"],
    "Code": [".py", ".js", ".html", ".css", ".json", ".sh"],
}

CONTENT_CATEGORIES = {
    "Financial": ["invoice", "bill", "receipt", "tax", "bank", "statement", "payment", "chequing", "banking", "account", "credit", "debit"],
    "Work": ["project", "report", "meeting", "resume", "proposal", "contract", "leads", "sales", "franchise"],
    "HR-Jobs": ["hr", "job", "hiring", "salary", "benefits", "employee"],
    "Legal": ["agreement", "legal", "court", "affidavit", "notary", "law", "signed"],
}

def extract_text(file_path: Path) -> str:
    """Extracts text from PDF, DOCX, or TXT."""
    ext = file_path.suffix.lower()
    text = ""
    try:
        if ext == ".txt":
            # Limit to first 100KB for safety
            with file_path.open(errors='ignore') as f:
                text = f.read(100_000)
        elif ext == ".pdf" and HAS_PDF:
            reader = PdfReader(file_path)
            meta = reader.metadata
            if meta:
                text += f" {meta.get('/Title', '')} {meta.get('/Subject', '')} "
            for i in range(min(5, len(reader.pages))):
                text += reader.pages[i].extract_text() or ""
        elif ext == ".docx" and HAS_DOCX:
            doc = Document(file_path)
            for i, para in enumerate(doc.paragraphs):
                if i > 50: break
                text += para.text + "\n"
    except Exception as e:
        logging.warning(f"Could not read {file_path.name}: {e}")
    return text.lower()

def classify(file_path: Path) -> Optional[str]:
    """Classifies documents into sub-folders."""
    content = f"{file_path.name} {extract_text(file_path)}".lower()
    scores = {cat: 0 for cat in CONTENT_CATEGORIES}
    for cat, keywords in CONTENT_CATEGORIES.items():
        for kw in keywords:
            if kw in content:
                scores[cat] += content.count(kw)
    if not any(scores.values()):
        return None
    return max(scores, key=lambda k: scores[k])

def get_safe_path(dest_dir: Path, filename: str) -> Path:
    """Handles duplicate filenames with counter (file (1).txt)."""
    dest_path = dest_dir / filename
    if not dest_path.exists():
        return dest_path
    
    path_obj = Path(filename)
    counter = 1
    while counter < 1000:
        new_name = f"{path_obj.stem} ({counter}){path_obj.suffix}"
        if not (dest_dir / new_name).exists():
            return dest_dir / new_name
        counter += 1
    return dest_dir / f"{path_obj.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{path_obj.suffix}"

def organize(folder: Path, dry_run: bool = True) -> tuple[int, int, int]:
    moved, skipped, errors = 0, 0, 0
    if not folder.exists():
        logging.warning(f"Folder not found: {folder}")
        return 0, 0, 0
    
    logging.info(f"--- Organizing: {folder} {'(DRY RUN)' if dry_run else ''} ---")
    
    # Create basic category folders
    if not dry_run:
        for cat in list(CATEGORIES.keys()) + ["Misc"]:
            (folder / cat).mkdir(exist_ok=True)

    # Use list() to avoid issues with modifying the directory while iterating
    for item in list(folder.iterdir()):
        if item.is_dir() or item.name.startswith('.'):
            continue

        # Find category
        ext = item.suffix.lower()
        target_cat = "Misc"
        for cat, extensions in CATEGORIES.items():
            if ext in extensions:
                target_cat = cat
                break
        
        # Handle Documents sub-folders
        dest_dir = folder / target_cat
        if target_cat == "Documents":
            sub_cat = classify(item)
            if sub_cat:
                dest_dir = dest_dir / sub_cat
            else:
                dest_dir = dest_dir / "Misc"
        
        # Ensure destination exists
        if not dry_run and not dest_dir.exists():
            dest_dir.mkdir(parents=True, exist_ok=True)

        # Move file safely
        final_dest = get_safe_path(dest_dir, item.name)
        try:
            if dry_run:
                logging.info(f"[DRY RUN] Would move: {item.name} -> {dest_dir.name}/{final_dest.name}")
            else:
                shutil.move(str(item), str(final_dest))
                logging.info(f"Moved: {item.name} -> {dest_dir.name}/{final_dest.name}")
            moved += 1
        except Exception as e:
            logging.error(f"Error moving {item.name}: {e}")
            errors += 1
            
    return moved, skipped, errors

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Organize files in Downloads and Documents.")
    parser.add_argument("--run", action="store_true", help="Execute the moves (disable dry-run)")
    args = parser.parse_args()

    dry_run = not args.run
    
    # Check dependencies
    if not HAS_PDF:
        logging.warning("⚠️  'pypdf' not installed. PDF content analysis disabled.")
    if not HAS_DOCX:
        logging.warning("⚠️  'python-docx' not installed. DOCX content analysis disabled.")

    if dry_run:
        print("\n🔍 DRY RUN MODE: No files will be moved. Use --run to execute.\n")

    total_moved = 0
    total_errors = 0

    for folder in FOLDERS_TO_ORGANIZE:
        m, s, e = organize(folder, dry_run=dry_run)
        total_moved += m
        total_errors += e
    
    print(f"\n--- Summary ---")
    print(f"Total Files Processed: {total_moved + total_errors}")
    print(f"Successful Moves: {total_moved}")
    print(f"Errors: {total_errors}")
    if dry_run:
        print("\n(This was a dry run. No changes were made.)")
    else:
        print("\nOrganization complete!")
