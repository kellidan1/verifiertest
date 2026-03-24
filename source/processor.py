from source.downloader import download_marksheet
from source.ocr_tesseract import extract_text_with_ocr
from source.ocr_paddle import extract_text_with_paddleocr
from source.verifier import verify_marksheet_data

def process_student_record(student_data):
    row, name, m1, link1, m2, link2, merit, percentage, h1, h2 = student_data

    def process_sem(link, marks, label, col_name):
        if not link:
            print(f"[{name}] Skipping {label}, no link provided")
            return None, None

        filename = f"{label}\\row{row}_{col_name}"
        file = download_marksheet(link, filename)

        if not file:
            return None, None

        ocr_text = extract_text_with_ocr(file)
        res = verify_marksheet_data(ocr_text, name, marks)

        need_name = not res["name_match"]
        need_marks = not res["marks_match"]

        if need_name or need_marks:
            ocr_text_p = extract_text_with_paddleocr(file)
            res_p = verify_marksheet_data(
                ocr_text_p,
                name if need_name else None,
                marks if need_marks else None
            )

            if need_name:
                res["name_match"] = res_p["name_match"]
            if need_marks:
                res["marks_match"] = res_p["marks_match"]
                res["ocr_marks"] = res_p["ocr_marks"]

        print(f"[{name}] {label} Name: {'YES' if res['name_match'] else 'NO'}")
        print(f"[{name}] {label} Marks: {'YES' if res['marks_match'] else 'NO'}")

        if not res["marks_match"]:
            print(f"Expected: {marks}, OCR: {res['ocr_marks']}")

        return res["name_match"], res["marks_match"]

    s1n, s1m = process_sem(link1, m1, "Sem1", h1)
    print("-" * 30)
    s2n, s2m = process_sem(link2, m2, "Sem2", h2)

    return row, name, s1n, s1m, s2n, s2m, merit, percentage