import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
import string
import ssl
from urllib3.poolmanager import PoolManager
from requests.adapters import HTTPAdapter

# class for weaker ssl
class UnsafeTLSAdapter(HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        context = ssl.create_default_context()
        context.set_ciphers("DEFAULT@SECLEVEL=1")
        kwargs['ssl_context'] = context
        return super().init_poolmanager(*args, **kwargs)

# session with lower ssl security level
session = requests.Session()
session.mount("https://", UnsafeTLSAdapter())

BASE_URL = "https://www.bip.cz/"
MAX_RETRIES = 3
RETRY_DELAY = 2
LETTERS = ["A", "B", "D", "F", "H", "CH", "J", "K", "L", "M", "N", "O", "P", "R", "S", "Š", "T", "Z", "Ž"]

def fetch_with_retry(url, max_retries=MAX_RETRIES, delay=RETRY_DELAY):
    for attempt in range(max_retries):
        try:
            response = session.get(url, timeout=10)
            if response.status_code == 200:
                return response
            else:
                print(f"[{attempt+1}] Chyba {response.status_code} při načítání {url}")
        except Exception as e:
            print(f"[{attempt+1}] Výjimka při načítání {url}: {e}")
        time.sleep(delay)
    return None

farnosti_data = []

# go through all letters from a to z
for letter in LETTERS:
    print(f"Zpracovávám písmeno: {letter}")
    list_url = f"{BASE_URL}/cs/katalog/farnosti?f.Key={letter}"
    list_response = fetch_with_retry(list_url)

    if not list_response:
        print(f"Písmeno {letter} přeskočeno (chyba při načítání).")
        continue

    list_soup = BeautifulSoup(list_response.text, "html.parser")
    farnost_links = list_soup.select("table.table-catalog tbody tr td:first-child a")

    for a_tag in farnost_links:
        nazev_farnosti = a_tag.text.strip()
        detail_url = BASE_URL + str(a_tag["href"])

        detail_response = fetch_with_retry(detail_url)
        if detail_response:
            detail_soup = BeautifulSoup(detail_response.text, "html.parser")

            # find all emails
            email_tags = detail_soup.select("a[href^='mailto:']")
            emails = [tag.get_text(strip=True) for tag in email_tags]
            emaily_spojene = ", ".join(emails)
        else:
            nazev = nazev_farnosti
            emaily_spojene = ""
            print(f"Nepodařilo se načíst detail farnosti: {detail_url}")

        farnosti_data.append((nazev_farnosti, emaily_spojene))
        time.sleep(0.5)

# excel
try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti Plzeň")
    else:
        ws.title = "Farnosti Plzeň"

    if ws:
        ws.append(["Název farnosti", "E-mail(y)"])
        for farnost, email in farnosti_data:
            ws.append([farnost, email])

        try:
            wb.save("farnosti_kontakty_plzen.xlsx")
            print("Hotovo!")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")