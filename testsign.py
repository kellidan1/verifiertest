import os
import torch
import fitz  # PyMuPDF
from PIL import Image
from transformers import AutoImageProcessor, AutoModelForObjectDetection

# Initialize model and processor once to avoid reloading for every call
try:
    PROCESSOR = AutoImageProcessor.from_pretrained("mdefrance/yolos-base-signature-detection")
    MODEL = AutoModelForObjectDetection.from_pretrained("mdefrance/yolos-base-signature-detection")
except Exception as e:
    print(f"Error loading model: {e}")
    PROCESSOR = None
    MODEL = None

def detect_signatures(file_path):
    """
    Detects and prints signatures present in a file (image/pdf) using YOLOS model.
    
    Args:
        file_path (str): The path to the image or PDF file.
    """
    if PROCESSOR is None or MODEL is None:
        print("Error: Model not loaded correctly.")
        return

    # Check if input file exists
    if not os.path.isfile(file_path):
        print(f"Error: Input file '{file_path}' not found.")
        return

    try:
        images = []
        if file_path.lower().endswith('.pdf'):
            # Convert PDF to images
            doc = fitz.open(file_path)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                images.append((img, f"Page {page_num + 1}"))
            doc.close()
        else:
            # Load as image
            img = Image.open(file_path).convert("RGB")
            images.append((img, "Original Image"))

        print(f"Signatures detected in {file_path}:")
        
        signatures_found_total = False
        for i, (image, label_str) in enumerate(images):
            inputs = PROCESSOR(images=image, return_tensors="pt")
            with torch.no_grad():
                outputs = MODEL(**inputs)

            # Post-processing
            target_sizes = torch.tensor([image.size[::-1]])
            results = PROCESSOR.post_process_object_detection(outputs, threshold=0.7, target_sizes=target_sizes)[0]

            if len(results["scores"]) > 0:
                signatures_found_total = True
                print(f"  [{label_str}]:")
                
                # Ensure outputs/signature directory exists
                output_dir = os.path.join("outputs", "signature")
                os.makedirs(output_dir, exist_ok=True)
                
                for idx, (score, label, box) in enumerate(zip(results["scores"], results["labels"], results["boxes"])):
                    box_list = [round(i, 2) for i in box.tolist()]
                    xmin, ymin, xmax, ymax = box.tolist()
                    
                    # Crop and save the signature image
                    sig_crop = image.crop((xmin, ymin, xmax, ymax))
                    base_name = os.path.basename(file_path).replace('.', '_')

                    save_path = os.path.join(output_dir, f"sig_{base_name}_{label_str.replace(' ', '')}_{idx+1}.png")
                    sig_crop.save(save_path)
                    
                    print(f"    - Signature {idx+1} found with confidence {score.item():.2f}")
                    print(f"      Saved to: {save_path}")
                    
                    # Extract text around the signature
                    try:
                        img_w, img_h = image.size
                        # Define region for label
                        padding_w = (xmax - xmin) * 0.5 
                        crop_left = max(0, xmin - padding_w)
                        crop_right = min(img_w, xmax + padding_w)
                        crop_top = max(0, ymax - 10) 
                        crop_bottom = min(img_h, ymax + max(60, (ymax - ymin) * 0.9))
                        
                        label_crop_img = image.crop((crop_left, crop_top, crop_right, crop_bottom))
                        
                        from PIL import ImageEnhance, ImageFilter
                        label_crop_gray = label_crop_img.convert('L')
                        label_crop_bin = ImageEnhance.Contrast(label_crop_gray).enhance(2.5)
                        label_crop_bin = label_crop_bin.filter(ImageFilter.SHARPEN)
                        
                        import pytesseract
                        import re
                        extracted_texts = []
                        for psm in [3, 6, 7]:
                            text = pytesseract.image_to_string(label_crop_bin, config=f'--oem 3 --psm {psm}').strip()
                            if text: extracted_texts.append(text)
                        
                        found_label = False
                        for text_result in extracted_texts:
                            lines = [line.strip() for line in text_result.split('\n') if line.strip()]
                            for l in lines:
                                if re.search('[a-zA-Z]{3,}', l):
                                    finalized_text = re.sub(r'\s+', ' ', l).strip()
                                    if len(finalized_text) < 30:
                                        #print(f"      Text underneath: \"{finalized_text}\"")
                                        found_label = True

                                        if finalized_text == "Secretary" or "Principal" or "Controller":
                                            print("VERIFIED")
                                        break
                            if found_label: break
                        
                        if not found_label:
                            print(f"      Text underneath: [Not detected]")
                    except Exception:
                        print(f"      Text underneath: [Error]")

        if not signatures_found_total:
            print("- No signatures detected in the entire file.")

        if not signatures_found_total:
            print("- No signatures detected in the entire file.")

    except FileNotFoundError as e:
        print(f"Error: A file or directory was not found during processing: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    # Example usage
    import sys
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
        detect_signatures(test_file)
    else:
        print("Usage: python testsign.py <path_to_file>")
