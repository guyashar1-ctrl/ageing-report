"""
file_utils.py - כלי עזר לקבצים
================================
חילוץ ZIP ואיתור BKMVDATA.TXT.
"""

import os
import zipfile
import tempfile
import shutil

from ageing_report.config.constants import BKMVDATA_FILENAMES


def extract_zip(uploaded_file):
    """
    חילוץ קובץ ZIP לתיקייה זמנית.

    Args:
        uploaded_file: קובץ מ-Streamlit (UploadedFile) או נתיב

    Returns:
        str: נתיב לתיקייה הזמנית
    """
    temp_dir = tempfile.mkdtemp(prefix="ageing_")

    if hasattr(uploaded_file, 'read'):
        zip_path = os.path.join(temp_dir, "upload.zip")
        with open(zip_path, 'wb') as f:
            f.write(uploaded_file.read())
        uploaded_file.seek(0)
    else:
        zip_path = uploaded_file

    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(temp_dir)

    return temp_dir


def find_bkmvdata(directory):
    """
    חיפוש קובץ BKMVDATA.TXT בתיקייה (כולל תת-תיקיות).

    Returns:
        str: נתיב לקובץ, או None
    """
    for root, _dirs, files in os.walk(directory):
        for name in BKMVDATA_FILENAMES:
            if name in files:
                return os.path.join(root, name)
    return None


def cleanup_temp_dir(path):
    """מחיקת תיקייה זמנית."""
    if path and os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
