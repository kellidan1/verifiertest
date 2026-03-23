import fitz
import numpy as np
from PIL import Image
import io
from source.config import paddle_ocr,BASE_DIR
import os


def extract_text_with_paddleocr(file_path):
    file_path = os.path.abspath(file_path)
    text = ""

    try:
        if file_path.lower().endswith('.pdf'):
            doc = fitz.open(file_path)
            for page in doc:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                results = paddle_ocr.ocr(np.array(img), cls=True)

                if results and results[0]:
                    for line in results[0]:
                        text += line[1][0] + " "
            doc.close()

        else:
            img = Image.open(file_path)
            results = paddle_ocr.ocr(np.array(img), cls=True)

            if results and results[0]:
                for line in results[0]:
                    text += line[1][0] + " "

    except Exception as e:
        print(f"PaddleOCR error: {e}")

    return text