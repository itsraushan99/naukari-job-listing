import json
import os
import shutil
from datetime import datetime, timedelta
from pathlib import Path

RESUME_DIR = Path(__file__).parent / "uploaded_resumes"
META_FILE = RESUME_DIR / "resume_meta.json"
EXPIRY_HOURS = 48


def _load_meta():
    if META_FILE.exists():
        try:
            with open(META_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return None


def _save_meta(filename, upload_time):
    RESUME_DIR.mkdir(exist_ok=True)
    with open(META_FILE, 'w') as f:
        json.dump({"filename": filename, "upload_time": upload_time.isoformat()}, f)


def save_resume(file_bytes, filename):
    """Save uploaded resume bytes to disk, replacing any existing one"""
    RESUME_DIR.mkdir(exist_ok=True)
    # Remove old resume files first
    delete_resume()
    dest = RESUME_DIR / filename
    with open(dest, 'wb') as f:
        f.write(file_bytes)
    _save_meta(filename, datetime.now())
    return dest


def load_resume():
    """
    Load resume from disk if it exists and hasn't expired.
    Returns (file_path, meta) or (None, None).
    Automatically deletes if expired.
    """
    meta = _load_meta()
    if not meta:
        return None, None

    upload_time = datetime.fromisoformat(meta["upload_time"])
    expiry_time = upload_time + timedelta(hours=EXPIRY_HOURS)

    if datetime.now() >= expiry_time:
        delete_resume()
        return None, None

    file_path = RESUME_DIR / meta["filename"]
    if not file_path.exists():
        delete_resume()
        return None, None

    return file_path, meta


def delete_resume():
    """Delete resume file and metadata"""
    meta = _load_meta()
    if meta:
        file_path = RESUME_DIR / meta["filename"]
        if file_path.exists():
            file_path.unlink()
    if META_FILE.exists():
        META_FILE.unlink()


def get_expiry_info():
    """
    Returns a dict with expiry details for display.
    Keys: upload_time, expiry_time, remaining_hours, remaining_minutes, expired
    Returns None if no resume on disk.
    """
    meta = _load_meta()
    if not meta:
        return None

    upload_time = datetime.fromisoformat(meta["upload_time"])
    expiry_time = upload_time + timedelta(hours=EXPIRY_HOURS)
    remaining = expiry_time - datetime.now()

    if remaining.total_seconds() <= 0:
        return {"expired": True}

    total_seconds = int(remaining.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes = remainder // 60

    return {
        "expired": False,
        "filename": meta["filename"],
        "upload_time": upload_time.strftime("%d %b %Y, %I:%M %p"),
        "expiry_time": expiry_time.strftime("%d %b %Y, %I:%M %p"),
        "remaining_hours": hours,
        "remaining_minutes": minutes,
    }
