import fitz
import pytesseract
from PIL import Image
import io

# Open the PDF
doc = fitz.open('tempmarksheet/KHUSHISHAIKH_marksheet.pdf')
page = doc.load_page(0)

# Get the pixmap (image) of the page
pix = page.get_pixmap(dpi=300) # Higher DPI for better OCR

# Convert pixmap to PIL Image
img = Image.open(io.BytesIO(pix.tobytes("png")))

# Perform OCR
text = pytesseract.image_to_string(img)
print("--- OCR TEXT ---")
print(text)
print("----------------")
