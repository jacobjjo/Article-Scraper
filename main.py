import requests
import os
import pandas as pd
from tqdm import tqdm
import re
from os import path
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import sys

excel = './RECORDS 1-1000.xlsx'

if len(sys.argv) < 2:
    print("Usage:")
    print("python script.py <excel_file>")
    sys.exit(1)

excel = sys.argv[1]

output_excel = sys.argv[2] if len(sys.argv) > 2 else "output.xlsx"
print(output_excel)

if not os.path.exists(excel):
    print(f"ERROR: File not found: {excel}")
    sys.exit(1)

print(f"Using Excel file: {excel}")

# Iterate through all dois in spreadsheet
# Download from unpaywall api
# Remove all valid pdfs from prospective doi list

# DEBUG
downloading = True

MAX_WORKERS = 16

session = requests.Session()

lock = threading.Lock()

def verify_pdf_bytes(content: bytes):
    return content[:5] == b"%PDF-"

def renamer(filename, doi):
    filename = re.sub(r'<[^>]+>', '', str(filename))
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)

    short = " ".join(filename.split(" ")[:4])

    doi_safe = doi.replace("/", "_")

    return f"{short}_{doi_safe}.pdf"

def extract_pdf_link(content: bytes):
    html = content.decode("utf-8", errors="ignore")
    soup = BeautifulSoup(html, "html.parser")

    for link in soup.find_all("a", href=True):
        href = link["href"]
        if ".pdf" in href.lower():
            return href

    return None

def detect_columns(df):

    detected = {}
    print(df.columns)
    for col in df.columns:

        clean = str(col).strip().lower()

        if "doi" in clean:
            detected["doi"] = col

        elif "title" in clean:
            detected["title"] = col

    return detected

try:
    columns = detect_columns(pd.read_excel(excel, header=10))
    df = pd.read_excel(
        excel,
        usecols=lambda c: str(c).strip().lower() in [
        "doi",
        "title"],
        header=10
    )
except pd.errors.ParserError as e:
    print(f"ERROR: Failed to parse Excel file: {e}")
    sys.exit(1)

articles = {}

for index, row in df.iterrows():

    doi = str(row['DOI']).strip()

    if doi == "nan":
        continue

    articles[doi] = {
            'doi': doi,
            'title': row['Title'],
            'url': "",
            'downloaded': False,
            'path': "",
            'row': index
    }

EMAIL = "jacobjeremiah42@gmail.com"

output_dir = "pdfs"

os.makedirs(output_dir, exist_ok=True)

print(f"Total DOIs to process: {len(articles)}")

# ---------------- MULTITHREADED DOWNLOAD FUNCTION ----------------

def download_pdf(article):

    try:

        url = f"https://api.unpaywall.org/v2/{article['doi']}?email={EMAIL}"

        r = session.get(url, timeout=15).json()

        pdf_url = None

        if r.get("best_oa_location"):

            pdf_url = r["best_oa_location"].get("url_for_pdf")

            article["url"] = r["best_oa_location"].get("url")

        if pdf_url:

            pdf_response = session.get(pdf_url, timeout=20)

            content = pdf_response.content

            if verify_pdf_bytes(content):

                filename = renamer(article['title'], article['doi'])

                file_path = os.path.join(output_dir, filename)

                with open(file_path, "wb") as f:
                    f.write(content)

                article["downloaded"] = True
                article["path"] = file_path

                return True

    except Exception as e:
        print(f"ERROR {article['doi']}: {e}")

    return False

# ---------------- MULTITHREADED EXECUTION ----------------

print("First pass of downloading PDFs from Unpaywall API")

futures = []

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

    for _, article in articles.items():

        futures.append(
            executor.submit(download_pdf, article)
        )

    for future in tqdm(
        as_completed(futures),
        total=len(futures),
        desc="Downloading PDFs"
    ):
        future.result()

# ---------------- COUNT REMAINING ----------------

remaining = [
    a for a in articles.values()
    if not a["downloaded"]
]

print(f"First pass complete. {len(remaining)} DOIs remain without valid PDFs.")

# ---------------- UPDATE EXCEL ----------------

for _, article in tqdm(articles.items(), desc="Editing Excel"):

    if article["downloaded"] == True:
        df.at[article['row'], 'PDF Downloaded? (Y/N)'] = "Y"
    else:
        df.at[article['row'], 'PDF Downloaded? (Y/N)'] = "N"

df.to_excel(
    output_excel,
    index=False,
    sheet_name="savedrecs"
)

print("Done.")