from sympy.matrices.kind import num_mat_mul
import fitz
import io
import numpy as np
from paddleocr import PaddleOCR
from word2number import w2n
from PIL import Image
import re

# Initialize PaddleOCR
# use_angle_cls=True enables text direction detection
# lang='en' specifies English language
ocr = PaddleOCR(use_angle_cls=True, lang='en')

def process_ocr(image):
    # Convert PIL Image to numpy array (RGB to BGR as PaddleOCR expects)
    img_array = np.array(image)
    
    # PaddleOCR usually works best with BGR if using cv2, 
    # but the ocr method handles numpy arrays.
    # If the image is RGB, we might want to convert if accuracy is low.
    
    print("--- Running OCR with PaddleOCR ---")
    results = ocr.ocr(img_array, cls=True)

    full_content = ""

    if results and results[0]:
        for line in results[0]:
            # result[0] contains a list of [box, (text, score)]
            text = line[1][0]
            full_content += text + " "

        full_content = re.sub(r' +', ' ', full_content)
        pattern = r'(?<=Rupees).*?(?=Only)'
        matches = re.findall(pattern, full_content, re.DOTALL)
        result = matches[0] if len(matches) == 1 else None
        number = w2n.word_to_num(result)
        if number<=300000:
            print(number)
        else:
            print("Not Eligible")
    else:
        print("No text detected.")




# Open the PDF
doc = fitz.open('test_data/SCAN0003.pdf')
page = doc.load_page(0)

# Get the pixmap (image) of the page
pix = page.get_pixmap(dpi=300) # Higher DPI for better OCR

# Convert pixmap to PIL Image
image = Image.open(io.BytesIO(pix.tobytes("png")))

# Process with OCR
process_ocr(image)

doc.close()