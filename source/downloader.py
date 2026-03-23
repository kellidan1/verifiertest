import os
import re
import requests
from source.config import BASE_DIR   # 👈 IMPORTANT

def download_marksheet(url, save_filename="marksheet"):
    """
    Downloads a file (image or pdf) from the given URL 
    and saves it into the main project's 'temp' directory.
    """

    if not url:
        print("No valid URL provided to download.")
        return None

    # ✅ Google Drive conversion (UNCHANGED - your original logic)
    if 'drive.google.com' in url:
        match_d = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
        if match_d:
            file_id = match_d.group(1)
            url = f"https://drive.google.com/uc?export=download&id={file_id}"
        else:
            match_id = re.search(r'[?&]id=([a-zA-Z0-9_-]+)', url)
            if match_id:
                file_id = match_id.group(1)
                url = f"https://drive.google.com/uc?export=download&id={file_id}"

    # ✅ FIX: correct temp folder path
    folder_path = os.path.join(BASE_DIR, "temp")

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    try:
        print(f"Downloading from {url} ...")
        response = requests.get(url, stream=True)
        response.raise_for_status()

        # ✅ SAME logic as your original (no changes)
        content_type = response.headers.get('content-type', '').lower()

        if 'application/pdf' in content_type:
            ext = 'pdf'
        elif 'image/jpeg' in content_type or 'image/jpg' in content_type:
            ext = 'jpg'
        elif 'image/png' in content_type:
            ext = 'png'
        else:
            ext = url.split('.')[-1][:4]
            if not ext.isalnum():
                ext = 'pdf'

        # ✅ FIX: use correct folder_path
        file_path = os.path.join(folder_path, f"{save_filename}.{ext}")

        with open(file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)

        print(f"Saved to {file_path}")
        return file_path

    except requests.exceptions.RequestException as e:
        print(f"Error downloading the file: {e}")
        return None