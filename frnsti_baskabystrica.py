import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time

BASE_URL = "https://schematizmus.bbdieceza.sk"
MAIN_URL = f"{BASE_URL}/dekanaty-a-farnosti"
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

# list for storage of final data to be printed
farnosti_data = []

# parsing the html code of the page
for dekanat_div in soup.select("div > h4.font-weight-bold"):
    parent_div = dekanat_div.parent
    if not parent_div:
        continue
    farnosti_links = parent_div.select("div.row div.col-3 a")
    
    for farnost_link in farnosti_links:
        farnost_url = farnost_link['href']
        # load parish detail
        detail_response = fetch_with_retry(farnost_url)
        if not detail_response:
            print(f"Nepodařilo se načíst detail farnosti: {farnost_url}")
            continue
        
        detail_soup = BeautifulSoup(detail_response.text, "html.parser")
        
        # parish name in h2
        nazev_farnosti_tag = detail_soup.select_one("div.card-body h2")
        nazev_farnosti = nazev_farnosti_tag.get_text(strip=True) if nazev_farnosti_tag else farnost_link.text.strip()
        
        # parish email
        email_tag = detail_soup.select_one("a[href^='mailto:']")
        email = email_tag.get_text(strip=True) if email_tag else ""
        
        farnosti_data.append((nazev_farnosti, email))
        print(f"Načteno: {nazev_farnosti} - {email}")
        time.sleep(0.5)

# save to excel
wb = Workbook()
if not wb.active:
    ws = wb.create_sheet("Farnosti bbdieceza")
else:
    ws = wb.active
    ws.title = "Farnosti bbdieceza"

if ws:
    ws.append(["Název farnosti", "E-mail"])
    for nazev, email in farnosti_data:
        ws.append([nazev, email])

    try:
        wb.save("farnosti_bbdieceza.xlsx")
        print("Hotovo!")
    except PermissionError:
        print("Nelze uložit soubor - možná je otevřený v jiném programu")
    except Exception as e:
        print(f"Chyba při ukládání souboru: {e}")
else:
    print("Nepodařilo se vytvořit list v Excelu")