# AutoKorea — Encar Scraper Pro

AutoKorea is a Windows desktop application for collecting and filtering used-car listings from Encar, a South Korean vehicle marketplace. Configure one or more search filters, run a bounded scrape, inspect the listings and progress log, and keep results in a local SQLite database and Excel workbook.

## What the app does

### Search and filtering

- Opens Encar's listing site in an embedded Qt WebEngine browser.
- Reads listing pages and uses Encar's paginated data endpoint for later result pages, avoiding reliance on a fragile page-number click.
- Starts with a conservative limit of 5 pages. The page control can be set from 1 to 50 pages; each page requests up to 200 listings.
- Supports multiple enabled filters. Each filter can specify a vehicle brand, model, extra title keywords, a year range, and month limits at the range endpoints.
- Requires all criteria inside a filter to match; any enabled filter can accept a listing. Disable, edit, remove, or clear filters in the interface.
- Extracts listing details such as the title, price, model year/month, mileage, fuel type, location, image URL, and Encar listing URL when available.
- Avoids inserting the same Encar car ID into the database more than once.

### Interface and language

- Presents download controls, filters, an embedded browser, results, progress, and a detailed scraping log in one PyQt desktop window.
- Defaults to Albanian and includes Albanian and British flag buttons for switching the interface and log between Albanian and English.
- Shows page progress and a brand breakdown as results are processed.

### Saved data

- Stores unique listings in `car_listings.db`, a local SQLite database with a `seen_cars` table.
- Creates a timestamped Excel workbook named `makinat_YYYY-MM-DD_HH-MM-SS.xlsx` for each run and appends the accepted listings to it.
- Keeps both outputs beside the application/source file. Existing database contents are retained between runs.

## Requirements

- Windows
- Python 3.9 or newer
- Internet access to load Encar and retrieve listing data
- Python packages listed in `requirements.txt`

Install dependencies from PowerShell in the project folder:

```powershell
python -m pip install -r requirements.txt
```

Run the application:

```powershell
python main.py
```

## Project files

- `main.py` — PyQt application, Encar browser and page workers, filters, parsing, deduplication, SQLite storage, Excel export, and language translations.
- `requirements.txt` — Python package dependencies.
- `car_listings.db` — included local SQLite snapshot; the application continues adding new unique listings here.
- `makinat_*.xlsx` — example exports from previous runs.
- `.gitignore` — excludes build environments, generated logs, bytecode, and executable files.

## Data and limitations

This tool collects listing information for personal research. Availability, prices, fields, and page structure are controlled by Encar and may change. Results reflect what the site returns during a run; check the original listing before making a purchase decision. Use the scraper responsibly and follow Encar's terms and applicable rules.

The compiled `main.exe` is not stored in this Git repository because it is approximately 137 MB, above GitHub's standard 100 MB per-file limit. The Python source can be run with the steps above; distribute the executable separately through a GitHub Release or another suitable file host if needed.
