import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

BASE_URL = "https://www.ado.cz"
MAIN_URL = f"{BASE_URL}/Katalog/Farnosti/"
MAX_RETRIES = 3
RETRY_DELAY = 2
farnosti_data = []

def fetch_with_retry(url, max_retries=MAX_RETRIES, delay=RETRY_DELAY):
    for attempt in range(max_retries):
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                return response
            else:
                print(f"[{attempt+1}] Chyba {response.status_code} při načítání {url}")
        except Exception as e:
            print(f"[{attempt+1}] Výjimka při načítání {url}: {e}")
        time.sleep(delay)
    return None

main_response = fetch_with_retry(MAIN_URL)

if not main_response:
    print("Nepodařilo se načíst hlavní stránku. Konec.")
    exit(1)

soup = BeautifulSoup(main_response.text, "html.parser")
table_rows = soup.select("table.format.w100 tr.kat_tr_1.trbg")

for row in table_rows:
    td = row.find("td")
    if not td:
        continue

    a_tag = td.find("a")
    if not a_tag:
        continue

    nazev_obce = a_tag.text.strip()
    nazev_farnosti = f"Římskokatolická farnost {nazev_obce}"
    detail_url = BASE_URL + str(a_tag["href"])

    # load parish detail
    detail_response = fetch_with_retry(detail_url)
    if detail_response:
        detail_soup = BeautifulSoup(detail_response.text, "html.parser")
        email = ""
        for p_tag in detail_soup.select("div.kontakty p"):
            strong = p_tag.find("strong")
            if strong and "E-mail" in strong.text:
                email = p_tag.get_text(strip=True).replace("E-mail:", "").strip()
                break
    else:
        email = ""
        print(f"Nepodařilo se načíst detail farnosti: {detail_url}")

    farnosti_data.append((nazev_farnosti, email))
    time.sleep(0.5)

try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti olomoucké diecéze")
    else:
        ws.title = "Farnosti olomoucké diecéze"

    if ws:
        ws.append(["Název farnosti", "E-mail"])
        for farnost, email in farnosti_data:
            ws.append([farnost, email])

        try:
            wb.save("farnosti_kontakty_olomouc.xlsx")
            print("Hotovo!")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")