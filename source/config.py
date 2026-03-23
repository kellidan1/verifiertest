from paddleocr import PaddleOCR
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
paddle_ocr = PaddleOCR(use_angle_cls=True, lang='en')