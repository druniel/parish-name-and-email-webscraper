from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
import time

# browser set up
options = Options()
options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# main url
base_url = "https://katalog.bcb.cz"
main_url = f"{base_url}/Katalog/Farnosti"

# load main website
driver.get(main_url)
time.sleep(2)

rows = driver.find_elements(By.CSS_SELECTOR, "table.format.w100 tr.trbg[class*='kat_tr_']")

farnosti = []

for i in range(len(rows)):
    try:
        # work with a new list of rows
        rows = driver.find_elements(By.CSS_SELECTOR, "table.format.w100 tr.trbg[class*='kat_tr_']")
        if i >= len(rows):
            print(f"Řádek {i} již neexistuje")
            continue
            
        row = rows[i]
        strong_tags = row.find_elements(By.TAG_NAME, "strong")
        if not strong_tags:
            continue
        nazev_farnosti = strong_tags[0].text.strip()
        
        try:
            odkaz = row.find_element(By.TAG_NAME, "a").get_attribute("href")
            if not odkaz:
                print(f"Prázdný odkaz pro farnost {nazev_farnosti}")
                continue
            if odkaz.startswith("/"):
                odkaz = base_url + odkaz
        except Exception as e:
            print(f"Nelze získat odkaz pro farnost {nazev_farnosti}: {e}")
            continue

        driver.get(odkaz)
        time.sleep(1)

        emaily = driver.find_elements(By.XPATH, "//a[starts-with(@href, 'mailto:')]")
        email_list = [e.text.strip() for e in emaily]

        farnosti.append((nazev_farnosti, email_list))

        # back to main url and reload all rows
        driver.get(main_url)
        time.sleep(1)
        rows = driver.find_elements(By.CSS_SELECTOR, "table.format.w100 tr.trbg[class*='kat_tr_']")

    except Exception as e:
        print(f"Chyba u řádku: {e}")
        continue

# close browser
driver.quit()

# save to excel
try:
    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Farnosti budějovice")
    else:
        ws.title = "Farnosti budějovice"

    ws.append(["Název farnosti", "E-mail 1", "E-mail 2", "E-mail 3", "E-mail 4"])

    for nazev, email_list in farnosti:
        # fill in void if smaller number of email per parish
        emails = email_list + [''] * (4 - len(email_list))
        ws.append([nazev] + emails[:4])  # just first 4 emails per parish

    try:
        wb.save("farnosti_budejovice.xlsx")
        print("Hotovo!")
    except PermissionError:
        print("Nelze uložit soubor - možná je otevřený v jiném programu")
    except Exception as e:
        print(f"Chyba při ukládání souboru: {e}")
except Exception as e:
    print(f"Chyba při práci s Excelem: {e}")