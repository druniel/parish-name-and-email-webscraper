import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

BASE_URL = "https://doo.cz"
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

main_url = f"{BASE_URL}/katalog/farnosti/"
main_response = fetch_with_retry(main_url)

if not main_response:
    print("Nepodařilo se načíst hlavní stránku. Konec.")
    exit(1)

soup = BeautifulSoup(main_response.text, "html.parser")
farnosti_data = []

for a_tag in soup.select("li a.link_text"):
    strong = a_tag.find("strong")
    if not strong:
        continue

    nazev_obce = strong.text.strip()
    nazev_farnosti = f"Římskokatolická farnost {nazev_obce}"
    detail_url = BASE_URL + str(a_tag["href"])

    # load detail with retry
    detail_response = fetch_with_retry(detail_url)
    if detail_response:
        detail_soup = BeautifulSoup(detail_response.text, "html.parser")
        email_tag = detail_soup.select_one("a[href^='mailto:']")
        email = email_tag.text.strip() if email_tag else ""
    else:
        email = ""
        print(f"Nepodařilo se načíst detail farnosti: {detail_url}")

    farnosti_data.append((nazev_farnosti, email))
    time.sleep(0.5)

try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti ostravská diecéze")
    else:
        ws.title = "Farnosti ostravská diecéze"

    if ws:
        ws.append(["Název farnosti", "E-mail"])
        for farnost, email in farnosti_data:
            ws.append([farnost, email])

        try:
            wb.save("farnosti_kontakty_ostrava.xlsx")
            print("Hotovo!")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")