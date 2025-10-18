import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

URL = "https://www.abu.sk/schematizmy/schematizmus-trnavskej-arcidiecezy-podla-nazvu-farnosti"

response = requests.get(URL)
if response.status_code != 200:
    raise SystemExit(f"Chyba při načítání stránky: HTTP {response.status_code}")

soup = BeautifulSoup(response.text, "html.parser")

farnosti_data = []

# find all contact blocks
for item in soup.select("div.contacts-item"):
    # parish name is in h4
    h4 = item.find("h4")
    if not h4:
        continue
    nazev = h4.get_text(separator=" ").strip()
    
    # email is in address
    email = ""
    for addr in item.find_all("address"):
        txt = addr.get_text()
        if "e‑mail:" in txt or "e-mail:" in txt:
            email = txt.split(":", 1)[1].strip()
            email = email.replace("(at)", "@").strip()
            break
    
    farnosti_data.append((nazev, email))

# excel
try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Trnavská arcidiecéza")
    else:
        ws.title = "Trnavská arcidiecéza"

    if ws:
        ws.append(["Název farnosti", "E-mail"])
        for nazev, email in farnosti_data:
            ws.append([nazev, email])

        try:
            wb.save("farnosti_trnava.xlsx")
            print("Hotovo! Výstup uložen jako 'farnosti_trnava.xlsx'")
        except PermissionError:
            print("Nelze uložit soubor - možná je otevřený v jiném programu")
        except Exception as e:
            print(f"Chyba při ukládání souboru: {e}")
    else:
        print("Nepodařilo se vytvořit list v Excelu")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")