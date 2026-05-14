import requests
import os
import pandas as pd
from tqdm import tqdm
import re
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import sys

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
    for col in df.columns:

        clean = str(col).strip().lower()

        if "doi" in clean:
            detected["doi"] = col

        elif "title" in clean:
            detected["title"] = col

    return detected

def process_file(excel, progress_callback, output_excel="output.xlsx", filter_year=None, stop_event=None):

    progress_callback(0, "Reading Excel...")

    try:
        df = pd.read_excel(
            excel,
            usecols=lambda c: str(c).strip().lower() in [
            "doi",
            "title",
            "publication year"],
            header=10
        )
    except pd.errors.ParserError as e:
        print(f"ERROR: Failed to parse Excel file: {e}")
        sys.exit(1)

    progress_callback(10, "Processing DOIs...")

    articles = {}

    total = len(df)

    for index, row in df.iterrows():
        print(total, index)

        doi = str(row['DOI']).strip()
        year = str(row['Publication Year']).strip()

        if doi == "nan":
            continue
        
        if filter_year and year not in filter_year:
            continue

        if stop_event.is_set():
            return

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

    futures = []

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:

        for _, article in articles.items():

            futures.append(
                executor.submit(download_pdf, article)
            )

        total = len(futures)
        for completed, future in enumerate(as_completed(futures), 1):
            
            if stop_event.is_set():
                executor.shutdown(
                    wait=False,
                    cancel_futures=True
                )
                return
    
            future.result()

            percent = 10 + int((completed / total) * 80)

            progress_callback(
                percent,
                f"Downloaded {completed} of {total} PDFs"
            )

    # ---------------- UPDATE EXCEL ----------------

    total_articles = len(articles.values())
    index = 0
    for article in articles.values():

        if article["downloaded"] == True:
            df.at[article['row'], 'PDF Downloaded? (Y/N)'] = "Y"
        else:
            df.at[article['row'], 'PDF Downloaded? (Y/N)'] = "N"

        percent = 90 + int((index / total_articles) * 10)

        progress_callback(
            percent,
            f"Updating spreadsheet ({index+1}/{total_articles})"
        )

        index += 1

    df.to_excel(
        output_excel,
        index=False,
        sheet_name="savedrecs"
    )
