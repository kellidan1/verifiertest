import re

def verify_marksheet_data(ocr_text, expected_name=None, expected_marks=None):
    results = {"name_match": None, "marks_match": None, "ocr_marks": None}

    def norm(s):
        return re.sub(r'\s+', ' ', str(s)).strip().upper()

    text = norm(ocr_text)

    if expected_name:
        results["name_match"] = norm(expected_name) in text

    if expected_marks:
        marks = str(expected_marks)
        if marks in text:
            results["marks_match"] = True
            results["ocr_marks"] = marks
        else:
            results["marks_match"] = False
            match = re.search(r'(TOTAL|GRAND TOTAL).*?(\d{2,4})', text)
            results["ocr_marks"] = match.group(2) if match else "Not Found"

    return results