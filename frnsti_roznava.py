import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

BASE_URL = "https://www.burv.sk"
MAIN_URL = f"{BASE_URL}/farnosti-a-filialky"
MAX_RETRIES = 3
RETRY_DELAY = 2

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

# find all parishes
farnosti_divs = soup.select("div.gal.item_A, div.gal.item_B, div.gal.item_C, div.gal.item_D, div.gal.item_E, div.gal.item_F, div.gal.item_G, div.gal.item_H, div.gal.item_CH, div.gal.item_I, div.gal.item_J, div.gal.item_K, div.gal.item_L, div.gal.item_M, div.gal.item_N, div.gal.item_O, div.gal.item_P, div.gal.item_Q, div.gal.item_R, div.gal.item_S, div.gal.item_T, div.gal.item_U, div.gal.item_V, div.gal.item_W, div.gal.item_X, div.gal.item_Y, div.gal.item_Z")

farnosti_data = []

for div in farnosti_divs:
    a_tag = div.select_one("a")
    if not a_tag:
        continue
    relative_url = a_tag["href"]
    relative_url_str = str(relative_url).strip("/")
    detail_url = f"{BASE_URL}/{relative_url_str}"

    detail_response = fetch_with_retry(detail_url)
    if not detail_response:
        print(f"Nepodařilo se načíst detail farnosti: {detail_url}")
        continue

    detail_soup = BeautifulSoup(detail_response.text, "html.parser")
    nazev_farnosti = detail_soup.select_one("h1")
    email_tag = detail_soup.select_one("a[href^='mailto:']")

    nazev = nazev_farnosti.text.strip() if nazev_farnosti else "Neznámá farnost"
    email = email_tag.text.strip() if email_tag else ""

    farnosti_data.append((nazev, email))
    print(f"Načteno: {nazev} - {email}")
    time.sleep(0.5)

# excel
try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti rožňavská diecéze")
    else:
        ws.title = "Farnosti rožňavská diecéze"

    if ws:
        ws.append(["Název farnosti", "E-mail"])
        for nazev, email in farnosti_data:
            ws.append([nazev, email])

        try:
            wb.save("farnosti_roznava.xlsx")
            print("Hotovo! Uloženo jako 'farnosti_roznava.xlsx'")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")