import csv
import io
import os
import re
import tempfile

from typing import List, Tuple

import pdfplumber
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pdf2image import convert_from_path
from PIL import Image, ImageOps
import pytesseract

# Define the Drive API scope and folder ID containing the PDFs
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
FOLDER_ID = os.environ.get('DRIVE_FOLDER_ID', 'YOUR_FOLDER_ID_HERE')
CREDENTIALS_FILE = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json')

# Column headers for the output CSV
CSV_COLUMNS = [
    '月日','伝票番号','証憑番号','借方科目コード','借方科目名','借方補助コード',
    '借方口座名','借方部門コード','借方部門名','借方課税区分','借方事業区分',
    '借方消費税額自動計算か否か','借方軽減税率か否か','借方税率','借方控除割合',
    '借方取引金額','借方消費税等','借方税抜き金額','貸方科目コード','貸方科目名',
    '貸方補助コード','貸方口座名','貸方部門コード','貸方部門名','貸方課税区分',
    '貸方事業区分','貸方消費税額自動計算か否か','貸方軽減税率か否か','貸方税率',
    '貸方控除割合','貸方取引金額','貸方消費税等','貸方税抜き金額','取引先コード',
    '取引先名','取引先の事業者登録番号','元帳摘要','実際の仕入れ年月日表示区分',
    '実際の仕入れ年月日１','実際の仕入れ年月日２','収支区分コード','収支区分名',
    '内訳区分コード','内訳区分名'
]

def get_drive_service() -> 'googleapiclient.discovery.Resource':
    """Authenticate and return a Drive service object."""
    credentials = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=credentials)

def list_pdf_files(service) -> List[dict]:
    """List PDF files in the specified Drive folder."""
    query = f"'{FOLDER_ID}' in parents and mimeType='application/pdf' and trashed=false"
    results = service.files().list(q=query, fields='files(id, name)').execute()
    return results.get('files', [])

def download_pdf(service, file_id: str, dest_path: str) -> None:
    """Download a PDF file from Drive."""
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(dest_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    fh.close()

def extract_text_from_pdf(pdf_path: str) -> str:
    """Try to extract text directly from a PDF."""
    text = ''
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + '\n'
    except Exception:
        pass
    return text

def pdf_to_images(pdf_path: str) -> List[Image.Image]:
    """Convert PDF pages to processed images for OCR."""
    images = convert_from_path(pdf_path, dpi=300)
    processed: List[Image.Image] = []
    for img in images:
        gray = ImageOps.grayscale(img)
        # enlarge image to improve OCR accuracy (300dpi equivalent already)
        enlarged = gray
        threshold = enlarged.point(lambda x: 0 if x < 128 else 255, '1')
        processed.append(threshold)
    return processed

def ocr_images(images: List[Image.Image]) -> str:
    """Run OCR on a list of images."""
    text = ''
    for img in images:
        text += pytesseract.image_to_string(img, lang='jpn+eng') + '\n'
    return text

def parse_info(text: str) -> Tuple[str, str, str]:
    """Parse date, amount, and partner name from OCR text."""
    date_match = re.search(r'\d{4}[/-]\d{1,2}[/-]\d{1,2}', text)
    amount_match = re.search(r'[0-9]{1,3}(?:,[0-9]{3})*', text)
    # naive approach for partner name: first non-empty line
    name = ''
    for line in text.splitlines():
        line = line.strip()
        if line:
            name = line
            break
    date = date_match.group(0) if date_match else ''
    amount = amount_match.group(0).replace(',', '') if amount_match else ''
    return date, amount, name

def build_csv_row(date: str, amount: str, partner: str) -> List[str]:
    """Create a CSV row with parsed data."""
    row = [''] * len(CSV_COLUMNS)
    row[0] = date.replace('-', '/') if date else ''  # 月日
    row[15] = amount  # 借方取引金額
    row[33] = partner  # 取引先名
    return row

def process_pdfs(service) -> None:
    """Main processing function."""
    files = list_pdf_files(service)
    with open('output.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(CSV_COLUMNS)
        for file in files:
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                download_pdf(service, file['id'], tmp.name)
                text = extract_text_from_pdf(tmp.name)
                if not text.strip():
                    images = pdf_to_images(tmp.name)
                    text = ocr_images(images)
                date, amount, partner = parse_info(text)
                writer.writerow(build_csv_row(date, amount, partner))
            os.unlink(tmp.name)

def main() -> None:
    service = get_drive_service()
    process_pdfs(service)

if __name__ == '__main__':
    main()
