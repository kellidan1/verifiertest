# Marksheet Verifier Tool

An automated tool to verify student marksheets against Excel data using OCR.

## Features
- **Automatic Marksheet Download**: Fetches files from URLs (Google Drive links are automatically converted to direct downloads).
- **OCR-based Verification**: Uses **Tesseract** and **PaddleOCR** to extract text from PDFs and images.
- **Dynamic Naming**: Downloaded files are saved predictably as `row[rownum]_[colname].[ext]` where `colname` is pulled from the Excel header.
- **Visual Feedback**: Automatically highlights matching data in GREEN and mismatches in RED within an output Excel file.
- **GUI Interface**: Simple Tkinter-based user interface for selecting files and monitoring progress.

## Excel Structure (testcl.xlsx)
The tool expects the following column structure in the input Excel file:

| Col # | Column Name | Usage |
|-------|-------------|-------|
| 2 | Full Name | Student name for verification |
| 6 | Merit Scholarship | Used for percentage-based verification |
| 14 | Marks of Sem 1 | Expected marks for Semester 1 |
| 17 | Marksheet1 File URL | Link to Semester 1 marksheet |
| 19 | Marks of Sem 2 | Expected marks for Semester 2 |
| 22 | Marksheet2 File URL | Link to Semester 2 marksheet |
| 26 | Percentage | Calculated percentage |

### All Column Headers in `testcl.xlsx`:
1. Timestamp
2. Full Name - Done
3. P.R.No.
4. Gender
5. Religion
6. Merit Scholarship - Not Done
7. Free Studenship - Not Done
8. Student Aid Fund - Not Done 
9. Degree & Subject
10. Year of Study 
11. Date of Joining - Not Done
12. University/College
13. Degree Name
14. Marks of Sem 1 - Done
15. Out of Sem 1
16. Entitlement Marks of sem 1
17. Marksheet1 File URL - Not Done
18. M1Marks equvalent Cert.
19. Marks of Sem 2 - Done
20. Out of Sem 2
21. Entitlement Marks Sem 2
22. Marksheet2 File URL - Not Done
23. M2Marks equivalent Cert
24. Grand Total Marks - Ongoing
25. Total Out Of
26. Percentage - Ongoing
27. Class
28. Backlogs
29. Fees Paid - Ongoing
30. Fee recipt - Ongoing
31. Guardian Name
32. Income
33. Income (Words)
34. Income Cert. Issue
35. Income Certificate File URL
36. Residence Cert no.
37. Residence Issued date
38. Residence certificate
39. Ration Card No
40. Ration Card
41. Address
42. Email
43. Mobile
44. Account Holder
45. Bank Name
46. Branch
47. Account No.
48. IFSC Code
49. Bank Document
50. Other Scholarship
51. Portal Scholarship Names
52. Other Agency
53. Funding Agency
54. Declaration

## Setup & Running

### Requirements
- Python 3.x
- Tesseract OCR (installed on system)
- Poppler (for PDF to Image conversion)

### Installation
```bash
pip install -r requirements.txt
```

### Usage
1. Run `python main.py`.
2. Browse and select your `.xlsx` file.
3. Click "Verify Data".
4. The output will be saved in the `outputs/` folder with color coding.
5. Temporary downloads are stored in the `temp/` folder.
