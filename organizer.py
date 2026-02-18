import os
import shutil
import logging
from pathlib import Path
from datetime import datetime
from typing import Optional

# Optional dependencies for content analysis
try:
    from pypdf import PdfReader
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
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
            text = file_path.read_text(errors='ignore')
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
    """Handles duplicate filenames with timestamps."""
    dest_path = dest_dir / filename
    if not dest_path.exists():
        return dest_path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path_obj = Path(filename)
    return dest_dir / f"{path_obj.stem}_{ts}{path_obj.suffix}"

def organize(folder: Path):
    if not folder.exists():
        return
    logging.info(f"--- Organizing: {folder} ---")
    
    # Create basic category folders
    for cat in list(CATEGORIES.keys()) + ["Misc"]:
        (folder / cat).mkdir(exist_ok=True)

    for item in folder.iterdir():
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
                dest_dir.mkdir(exist_ok=True)
            else:
                dest_dir = dest_dir / "Misc"
                dest_dir.mkdir(exist_ok=True)

        # Move file safely
        final_dest = get_safe_path(dest_dir, item.name)
        try:
            shutil.move(str(item), str(final_dest))
            logging.info(f"Moved: {item.name} -> {dest_dir.name}/")
        except Exception as e:
            logging.error(f"Error moving {item.name}: {e}")

if __name__ == "__main__":
    for folder in FOLDERS_TO_ORGANIZE:
        organize(folder)
    print("\nOrganization complete!")
