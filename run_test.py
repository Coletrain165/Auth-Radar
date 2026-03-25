import pytesseract, sys, os
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\ColeMcComas\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
sys.path.insert(0, '.')
from auth_extractor import PDFExtractor
extractor = PDFExtractor()

with open('test_results.txt', 'w') as f:
    pdfs = [
        'Test Batch/William Cross Unskilled Auth 08052025.pdf',
        'Test Batch/Reina Figueroa Skilled Auth 06082025.pdf',
        'Test Batch/Taag Olson Escort Assistance 01142026.pdf',
        'Test Batch/Pedro Hernandez Skilled Auth 11102025-encrypted.pdf',
    ]
    fields = ['Patient Name', 'Auth #', 'Date Approved', 'Date Auth Expire', 'Patient ID']
    for pdf in pdfs:
        f.write(f'\n==== {os.path.basename(pdf)}\n')
        result = extractor.process_pdf(pdf)
        for k in fields:
            method = result.get(k + '_method', '')
            f.write(f'  {k}: {result.get(k, "NOT FOUND")}  [{method}]\n')

print("Done - see test_results.txt")
