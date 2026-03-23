import fitz
import pytesseract
from PIL import Image
import io
from source.config import BASE_DIR
import os

def extract_text_with_ocr(file_path):
    file_path = os.path.abspath(file_path)
    text = ""

    try:
        if file_path.lower().endswith('.pdf'):
            doc = fitz.open(file_path)
            for page in doc:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                text += pytesseract.image_to_string(img)
            doc.close()

        else:
            img = Image.open(file_path)
            text += pytesseract.image_to_string(img)

    except Exception as e:
        print(f"OCR error: {e}")

    return text