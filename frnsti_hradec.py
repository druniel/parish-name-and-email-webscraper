import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
import string

BASE_URL = "https://www.bihk.cz"
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

farnosti_data = []

# Prochází všechna písmena A–Z
for letter in string.ascii_uppercase:
    print(f"Zpracovávám písmeno: {letter}")
    list_url = f"{BASE_URL}/dieceze/diecezni-katalog/farnosti-filter/{letter}"
    list_response = fetch_with_retry(list_url)

    if not list_response:
        print(f"Písmeno {letter} přeskočeno (chyba při načítání).")
        continue

    list_soup = BeautifulSoup(list_response.text, "html.parser")
    farnost_links = list_soup.select("div.result-items li a")

    for a_tag in farnost_links:
        nazev_farnosti = a_tag.text.strip()
        try:
            if not a_tag:
                print(f"Přeskakuji farnost - chybějící odkaz")
                continue
            
            href = a_tag.get("href")
            if not href:
                print(f"Přeskakuji farnost - prázdný odkaz")
                continue
            
            detail_url = BASE_URL + str(href)

        except AttributeError as e:
            print(f"Chyba při zpracování odkazu: {e}")
            continue

        detail_response = fetch_with_retry(detail_url)
        if detail_response:
            detail_soup = BeautifulSoup(detail_response.text, "html.parser")

            # parish name
            try:
                nazev_tag = detail_soup.select_one("div.region-page-title h1")
                nazev = nazev_tag.get_text(strip=True) if nazev_tag else nazev_farnosti

            except AttributeError as e:
                print(f"Chyba při zpracování detailu farnosti {nazev_farnosti}: {e}")
                nazev = nazev_farnosti

            # find all emails
            email_tags = detail_soup.select("a[href^='mailto:']")
            emails = [tag.get_text(strip=True) for tag in email_tags]
            emaily_spojene = ", ".join(emails)
        else:
            nazev = nazev_farnosti
            emaily_spojene = ""
            print(f"Nepodařilo se načíst detail farnosti: {detail_url}")

        farnosti_data.append((nazev, emaily_spojene))
        time.sleep(0.5)

# save to excel
try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti Hradec")
    else:
        ws.title = "Farnosti Hradec"

    if ws:
        ws.append(["Název farnosti", "E-mail(y)"])
        for farnost, email in farnosti_data:
            ws.append([farnost, email])

        try:
            wb.save("farnosti_kontakty_hradec.xlsx")
            print("Hotovo!")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")