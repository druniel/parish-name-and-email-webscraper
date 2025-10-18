# Parish name and email webscraper

Simple collection of small scrapers that extract parish names and email addresses from diocesan catalog websites and save results to Excel (OpenPyXL).

## Repository layout
- One script per diocese (e.g. `frnsti_nitra.py`, `frnsti_praha.py`, `frnsti_brno.py`)
- Each script scrapes list/detail pages and writes an `.xlsx` file to the project root.

## Requirements
- Python 3.8+
- A requirements file is included: `requirements.txt`

Install dependencies from the included requirements file:
```bash
python -m pip install -r requirements.txt
```

Typical packages used (already in requirements.txt):
- requests
- beautifulsoup4
- openpyxl
- selenium (only for scripts that use a browser)
- webdriver-manager
- urllib3

## Usage
Run a script from the project directory (Windows):
```powershell
python frnsti_nitra.py
```
Replace the filename to run a different scraper.

Notes:
- Some scripts use Selenium (headless Chrome). `webdriver-manager` usually downloads a compatible ChromeDriver automatically.
- Respect remote servers: keep or increase delays between requests if needed.

## Output
- XLSX files saved in the project root, e.g. `farnosti_nitra.xlsx`
- Columns: typically `Název farnosti`, `E-mail` (scripts may vary)
