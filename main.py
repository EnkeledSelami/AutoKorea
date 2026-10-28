#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os

os.environ["QTWEBENGINE_DISABLE_SANDBOX"] = "1"
os.environ["QTWEBENGINE_CHROMIUM_FLAGS"] = "--disable-gpu"

import json, re, sqlite3, sys, datetime
import traceback
from pathlib import Path
import pandas as pd
from PyQt5.QtCore import QUrl, QTimer, Qt, pyqtSignal, QThread
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QSpinBox,
    QTextEdit, QListWidget, QListWidgetItem, QProgressBar, QSplitter, QGroupBox, QDialog,
    QLineEdit, QDialogButtonBox, QCheckBox, QTabWidget, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QComboBox, QMessageBox, QGridLayout
)
from PyQt5.QtGui import QColor, QFont, QPainter, QPen, QBrush, QPolygon, QIcon, QPixmap
from PyQt5.QtCore import QPoint, QSize
from PyQt5.QtWebEngineWidgets import QWebEngineView
from bs4 import BeautifulSoup
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen

APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
NOW = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
EXCEL_FILE = str(APP_DIR / f"makinat_{NOW}.xlsx")
DB_PATH = str(APP_DIR / "car_listings.db")
MAX_PAGES = 5
CURRENT_YEAR = datetime.date.today().year
MAX_YEAR = CURRENT_YEAR + 1  # was hard-coded 2025 -> silently dropped every 2026 car

LANGUAGE = "sq"

TRANSLATIONS = {
    "Shto/Ndrysho Filtër": "Add/Edit Filter", "Marka & Modeli": "Brand & Model",
    "Çdo Markë": "Any Brand", "Çdo Model": "Any Model", "Marka:": "Brand:",
    "Modeli:": "Model:", "Fjalë kyçe shtesë": "Additional keywords",
    "të ndara me presje": "comma separated", "Intervali i viteve": "Year range",
    "Pa kufizim": "Unrestricted", "Çdo": "Any", "Viti min:": "Min year:",
    "Viti max:": "Max year:", "Muaji min:": "Min month:", "Muaji max:": "Max month:",
    "Aktivizo këtë filtër": "Enable this filter", "Krijoni filtra të saktë për kërkimin e makinave. Filloni duke shtuar një filtër më poshtë.": "Create precise filters for car searches. Start by adding a filter below.",
    "Marka": "Brand", "Modeli": "Model", "Fjalë kyçe": "Keywords", "Vitet": "Years", "Muajt": "Months",
    "➕ Shto": "➕ Add", "✏️ Ndrysho": "✏️ Edit", "🗑️ Fshi": "🗑️ Delete", "🧹 Pastro": "🧹 Clear",
    "Asnjë përzgjedhje": "Nothing selected", "Ju lutem zgjidhni një filtër për ta ndryshuar.": "Please select a filter to edit.",
    "Ju lutem zgjidhni një filtër për ta fshirë.": "Please select a filter to delete.", "Konfirmo": "Confirm",
    "Fshi filtrin?": "Delete this filter?", "Pastro të gjitha?": "Clear all filters?",
    "Encar Scraper Pro – Filtro PËRPARA se të ruash": "Encar Scraper Pro – Filter before saving",
    "🎛️ Kontrollet e shkarkimit": "🎛️ Download controls", "Faqet:": "Pages:",
    "🚀 Fillo shkarkimin": "🚀 Start scraping", "⏳ Duke shkarkuar...": "⏳ Scraping...",
    "Gati – konfiguro filtrat fillimisht": "Ready – configure filters first", "Duke shkarkuar …": "Scraping …",
    "🔍 Filtrat e mençur": "🔍 Smart filters", "🌐 Shfletuesi": "🌐 Browser",
    "📊 Rezultatet": "📊 Results", "📝 Ditari i shkarkimit:": "📝 Scraping log:",
    "Shkarkim i përfunduar.": "Scraping complete.", "makina u ruajtën.": "cars saved.",
    "Përfunduar": "Finished", "Çmimi i padisponueshëm": "Price unavailable",
    "🚀 Filloi shkarkimi i ri": "🚀 New scraping run started",
    "🚀 Fillo shkarkimin": "🚀 Start scraping",
    "Asnjë përzgjedhje": "Nothing selected",
}

LOG_TRANSLATIONS = {
    "🌐 Duke ngarkuar faqen": "🌐 Loading page",
    "🚀 Filloi shkarkimi i ri": "🚀 New scraping run started",
    "Faqja ": "Page ",
    "faqen ": "page ",
    "Duke kaluar te faqja ": "Moving to page ",
    "Duke kërkuar faqen ": "Requesting page ",
    "Kontrolli i faqes tjetër u aktivizua: ": "Next-page control clicked: ",
    "Clicki i faqes tjetër nuk u konfirmua: ": "Next-page click was not confirmed: ",
    "Navigimi te faqja ": "Navigation to page ",
    " dështoi ": " failed ",
    "Po riprovohet...": "Retrying...",
    "Nuk u arrit faqja ": "Could not reach page ",
    " pas 3 përpjekjeve. ": " after 3 attempts. ",
    "nga kontrolli i faqes.": "using the page navigation control.",
    "Nuk u gjet kontrolli i faqes tjetër: ": "Could not find the next-page control: ",
    "nuk dha rezultate të reja; ndalova që të mos përsëris faqen 1.": "did not show new results; stopped to avoid repeating page 1.",
    "nuk u ngarkua; po provohet leximi i saj.": "failed to load; trying to read it anyway.",
    "të dhënat nuk u shfaqën pas 30 sekondash.": "listing data did not appear after 30 seconds.",
    "Parser: ": "Parsed: ", " makina (": " cars (", " makina, ": " cars, ",
    "⚠️ 0 makina. HTML u ruajt te ": "⚠️ 0 cars. HTML saved to ", " | klasat: ": " | classes: ",
    " makina u ruajtën pas filtrimit": " cars saved after filtering",
    "Markat në faqe: ": "Brands on page: ", "Gabim gjatë procesimit: ": "Processing error: ",
    "Year: ": "Year: ", "Passes: ": "Passes: ", "makina u ruajtën": "cars saved",
}


def translate_log(text, language):
    pairs = list(LOG_TRANSLATIONS.items())
    if language == "sq":
        pairs = [(english, albanian) for albanian, english in pairs]
    for source, target in sorted(pairs, key=lambda item: len(item[0]), reverse=True):
        text = text.replace(source, target)
    return text


def flag_icon(country):
    """Draw crisp flag icons because Qt 5's bundled font renders flag emoji as AL/GB text."""
    pixmap = QPixmap(42, 28)
    pixmap.fill(Qt.transparent)
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    if country == "gb":
        painter.fillRect(0, 0, 42, 28, QColor("#012169"))
        painter.setPen(QPen(Qt.white, 8, Qt.SolidLine, Qt.SquareCap))
        painter.drawLine(0, 0, 42, 28)
        painter.drawLine(42, 0, 0, 28)
        painter.setPen(QPen(QColor("#C8102E"), 3, Qt.SolidLine, Qt.SquareCap))
        painter.drawLine(0, 0, 42, 28)
        painter.drawLine(42, 0, 0, 28)
        painter.fillRect(16, 0, 10, 28, Qt.white)
        painter.fillRect(0, 9, 42, 10, Qt.white)
        painter.fillRect(19, 0, 4, 28, QColor("#C8102E"))
        painter.fillRect(0, 12, 42, 4, QColor("#C8102E"))
    else:
        painter.fillRect(0, 0, 42, 28, QColor("#E41E20"))
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(Qt.black))
        # Compact double-headed eagle silhouette: wings, twin necks/heads, body and talons.
        painter.drawPolygon(QPolygon([QPoint(21, 12), QPoint(17, 7), QPoint(12, 5), QPoint(8, 3),
                                      QPoint(9, 7), QPoint(5, 5), QPoint(7, 10), QPoint(3, 9),
                                      QPoint(7, 14), QPoint(13, 15), QPoint(18, 13)]))
        painter.drawPolygon(QPolygon([QPoint(21, 12), QPoint(25, 7), QPoint(30, 5), QPoint(34, 3),
                                      QPoint(33, 7), QPoint(37, 5), QPoint(35, 10), QPoint(39, 9),
                                      QPoint(35, 14), QPoint(29, 15), QPoint(24, 13)]))
        painter.drawPolygon(QPolygon([QPoint(18, 11), QPoint(21, 8), QPoint(24, 11), QPoint(23, 19),
                                      QPoint(27, 22), QPoint(23, 21), QPoint(21, 24), QPoint(19, 21),
                                      QPoint(15, 22), QPoint(19, 19)]))
        painter.drawPolygon(QPolygon([QPoint(18, 10), QPoint(14, 7), QPoint(12, 8), QPoint(14, 10), QPoint(17, 11)]))
        painter.drawPolygon(QPolygon([QPoint(24, 10), QPoint(28, 7), QPoint(30, 8), QPoint(28, 10), QPoint(25, 11)]))
        painter.drawPolygon(QPolygon([QPoint(19, 19), QPoint(15, 22), QPoint(19, 21)]))
        painter.drawPolygon(QPolygon([QPoint(23, 19), QPoint(27, 22), QPoint(23, 21)]))
    painter.end()
    return QIcon(pixmap)


def tr(text):
    return TRANSLATIONS.get(text, text) if LANGUAGE == "en" else text

LIST_BASE = "https://car.encar.com/list/car"
DETAIL_URL = "https://fem.encar.com/cars/detail/{}"


def build_list_url(page: int) -> str:
    """Encar's list page now takes a JSON `search` param (old ?sort=ModifiedDate is ignored)."""
    search = {
        "type": "car",
        "action": "(And.Hidden.N._.CarType.A.)",   # Encar's canonical public-listing action
        "title": "",
        "toggle": {},
        "layer": "",
        "sort": "MobileModifiedDate",               # newest first
    }
    raw = json.dumps(search, separators=(",", ":"), ensure_ascii=False)
    return f"{LIST_BASE}?page={page}&search={quote(raw, safe='')}"

# ---------- Konstantat të KORRIGJUARA ----------
CAR_BRANDS = {
    "Mercedes-Benz": ["벤츠", "메르세데스", "mercedes", "benz"],
    "BMW": ["BMW", "bmw", "비엠더블유"],
    "Audi": ["아우디", "audi"],
    "Volkswagen": ["폭스바겐", "폴크스바겐", "volkswagen"],
    "Hyundai": ["현대", "hyundai"],  # KORRIGJUAR
    "Kia": ["기아", "kia"],
    "Genesis": ["제네시스", "genesis"],
    "Toyota": ["토요타", "toyota"],
    "Honda": ["혼다", "honda"],
    "Nissan": ["닛산", "nissan"]
}

CAR_MODELS = {
    "Mercedes-Benz": ["C-Class", "E-Class", "S-Class", "GLE", "GLC", "A-Class", "C클래스", "E클래스", "S클래스"],
    "BMW": ["3 Series", "5 Series", "7 Series", "X3", "X5", "X1", "3시리즈", "5시리즈", "7시리즈"],
    "Audi": ["A3", "A4", "A6", "Q3", "Q5", "Q7"],
    "Volkswagen": ["Golf", "Tiguan", "Passat", "Jetta", "Touareg", "골프", "티구안"],
    "Hyundai": ["Sonata", "Elantra", "Tucson", "Santa Fe", "i30", "Avante", "쏘나타", "엘란트라", "투싼", "싼타페", "아반떼", "그랜저"],
    "Kia": ["Sorento", "Sportage", "Forte", "Optima", "Soul", "쏘렌토", "스포티지", "K5", "K3"],
    "Genesis": ["G70", "G80", "G90", "GV70", "GV80"],
    "Toyota": ["Camry", "Corolla", "RAV4", "Prius", "Highlander"],
    "Honda": ["Accord", "Civic", "CR-V", "Pilot", "Fit"],
    "Nissan": ["Altima", "Sentra", "Rogue", "Pathfinder", "Maxima"]
}

MODEL_KOREAN = {
    "C-Class": ["c클래스", "c-클래스"],
    "E-Class": ["e클래스", "e-클래스"],
    "S-Class": ["s클래스", "s-클래스"],
    "3 Series": ["3시리즈"],
    "5 Series": ["5시리즈"],
    "7 Series": ["7시리즈"],
    "A3": ["a3"], "A4": ["a4"], "A6": ["a6"],
    "Golf": ["골프"], "Tiguan": ["티구안"],
    "Sonata": ["쏘나타"], "Elantra": ["엘란트라"],
    "Tucson": ["투싼"], "Santa Fe": ["싼타페"],
    "i30": ["i30"], "Avante": ["아반떼"]
}

# ---------- Funksione ndihmëse - KORRIGJUAR ----------
import re


def extract_year_month(txt: str):
    """Ekstrakton vitin dhe muajin nga formatet Korean të ndryshme"""
    if not txt or txt.strip() == "N/A":
        return None, None

    txt = re.sub(r"\s+", " ", txt.strip())

    # PATTERNS TË PËRMIRËSUARA për formatet Korean
    patterns = [
        # Format: "23/08식", "22/05식(23년형)", "18/12식(19년형)"
        r"(\d{2,4})[./](\d{1,2})식",

        # Format: "23/08식(24년형)" - merr vitin e parë
        r"(\d{2,4})[./](\d{1,2})식\((\d{2,4})년형\)",

        # Format: "2023년 8월"
        r"(\d{2,4})년\s*(\d{1,2})월",

        # Format: "23년 08월"
        r"(\d{2,4})년\s*(\d{1,2})월",

        # Format: "2023 08"
        r"(\d{4})\s+(\d{1,2})",

        # Format vetëm viti: "2023식"
        r"(\d{2,4})식",

        # Format: "23식"
        r"(\d{2})식",
    ]

    for pattern in patterns:
        match = re.search(pattern, txt)
        if match:
            groups = match.groups()
            year_str = groups[0]
            month_str = groups[1] if len(groups) > 1 else "1"  # Default muaj

            try:
                year = int(year_str)
                month = int(month_str) if month_str.isdigit() else 1

                # Konverto 2-digit years në 4-digit
                if year < 100:
                    if year > 50:
                        year += 1900  # 51-99 -> 1951-1999
                    else:
                        year += 2000  # 00-50 -> 2000-2050

                # Valido muajin
                if not (1 <= month <= 12):
                    month = 1

                return year, month

            except ValueError:
                continue

    # FALLBACK: Provo të gjesh çdo numër 2-4 shifror si vit
    year_match = re.search(r"(\d{2,4})", txt)
    if year_match:
        try:
            year = int(year_match.group(1))
            if year < 100:
                year += 1900 if year > 50 else 2000
            # Kontrollo nëse është vit i arsyeshëm
            if 1990 <= year <= MAX_YEAR:
                return year, 1  # Default muaj
        except ValueError:
            pass

    return None, None


def test_year_extraction():
    """Test funksioni për të kontrolluar ekstraktimin e viteve"""
    test_cases = [
        "23/08식",
        "22/05식(23년형)",
        "18/12식(19년형)",
        "21/04식",
        "17/08식(18년형)",
        "2023년 8월",
        "20/11식",
        "14/10식(15년형)"
    ]

    print("=== TEST YEAR EXTRACTION ===")
    for case in test_cases:
        year, month = extract_year_month(case)
        print(f"'{case}' -> Year: {year}, Month: {month}")
    print("=============================\n")


# KORRIGUAR - Filtri më i saktë me DEBUG
def matches_filters_strict(title, year_txt, filters):
    if not filters:
        return True  # pa filtra → prano të gjitha

    title_lower = title.lower()
    year, month = extract_year_month(year_txt)

    # DEBUG: Print year extraction
    # print(f"DEBUG: '{year_txt}' -> Year: {year}, Month: {month}")

    for rule in filters:
        if not rule.get("enabled", True):
            continue

        # Kontrollo brand
        brand = rule.get("brand")
        brand_match = True  # Default të jetë True nëse nuk ka brand filter
        if brand and brand in CAR_BRANDS:
            brand_keywords = CAR_BRANDS[brand]
            brand_match = any(k.lower() in title_lower for k in brand_keywords)
            if not brand_match:
                continue  # Nëse nuk përputhet brand-i, provoj rregullin tjetër

        # Kontrollo model
        model = rule.get("model")
        model_match = True  # Default të jetë True nëse nuk ka model filter
        if model:
            if model in MODEL_KOREAN:
                model_keywords = MODEL_KOREAN[model]
                model_match = any(k.lower() in title_lower for k in model_keywords)
            else:
                model_match = model.lower() in title_lower
            if not model_match:
                continue

        # Kontrollo keywords
        keywords = rule.get("keywords", [])
        keyword_match = True  # Default të jetë True nëse nuk ka keywords
        if keywords:
            keyword_match = all(k.lower() in title_lower for k in keywords)
            if not keyword_match:
                continue

        # KONTROLLO VITET - më i saktë
        year_match = True
        if year is not None:  # Vetëm nëse kemi ekstraktuar vitin
            min_y = rule.get("min_year")
            max_y = rule.get("max_year")

            if min_y and year < min_y:
                year_match = False
            if max_y and year > max_y:
                year_match = False

            if not year_match:
                continue

        # KONTROLLO MUAJT - vetëm nëse kemi të dy year dhe month
        month_match = True
        if year is not None and month is not None:
            min_m = rule.get("min_month")
            max_m = rule.get("max_month")
            min_y = rule.get("min_year")
            max_y = rule.get("max_year")

            # Kontrollo muajin vetëm për vitet limit
            if min_m and min_y and year == min_y and month < min_m:
                month_match = False
            if max_m and max_y and year == max_y and month > max_m:
                month_match = False

            if not month_match:
                continue

        # Nëse të gjitha kontrollet kalojnë, filtri përputhet
        return True

    return False  # Asnjë filtër nuk përputhet


def save_excel(car):
    try:
        df = pd.DataFrame([car])
        file = Path(EXCEL_FILE)
        if file.exists():
            with pd.ExcelWriter(file, mode="a", engine="openpyxl", if_sheet_exists="overlay") as writer:
                df.to_excel(writer, index=False, header=False,
                            startrow=writer.sheets["Sheet1"].max_row)
        else:
            df.to_excel(file, index=False, header=True)
    except Exception as e:
        print(f"[Excel error]: {e}")


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""CREATE TABLE IF NOT EXISTS seen_cars (
        car_id TEXT PRIMARY KEY,
        url TEXT, title TEXT, price TEXT,
        mileage TEXT, year TEXT,
        first_seen TEXT, image_url TEXT
    )""")
    return conn



# ---------- HTML parsing (independent of Encar's hashed CSS-module suffixes) ----------
YEAR_RE = re.compile(r"\d{2,4}\s*[/.]\s*\d{1,2}\s*식|\d{4}\s*년")
DETAIL_ID_RE = re.compile(r"(?:/cars/detail/|carid=)(\d+)", re.I)
PRICE_RE = re.compile(r"[\d,]+\s*만원")


def _cls(name):
    """Match `ItemBigImage_<name>__<hash>` whatever the hash currently is."""
    return f'[class*="ItemBigImage_{name}__"]'


def _pick_image(item):
    img = item.select_one(_cls("thumb")) or item.select_one("img")
    if not img:
        return ""
    cands = [img.get("data-src", ""), img.get("src", "")]
    srcset = img.get("srcset", "")
    if srcset:
        cands.append(srcset.split(",")[0].strip().split(" ")[0])
    for c in cands:
        if c and not c.endswith("trans.gif") and not c.startswith("data:"):
            return "https:" + c if c.startswith("//") else c
    return ""


def _parse_item(item):
    # --- car id: enlog JSON, else the detail link ---
    car_id = None
    holder = item if item.has_attr("data-enlog-dt-param") else item.select_one("[data-enlog-dt-param]")
    if holder:
        try:
            car_id = json.loads(holder.get("data-enlog-dt-param", "{}")).get("carId")
        except Exception:
            pass
    if not car_id:
        a = item if item.name == "a" else item.select_one('a[href*="/cars/detail/"], a[href*="carid="]')
        m = DETAIL_ID_RE.search(a.get("href", "")) if a else None
        car_id = m.group(1) if m else None
    if not car_id:
        return None
    car_id = str(car_id)

    # --- title ---
    t = item.select_one(_cls("name")) or item.select_one("strong")
    title = t.get_text(" ", strip=True) if t else "N/A"

    # --- price ---
    price = "N/A"
    status = item.select_one(_cls("price_txt"))
    if status and status.get_text(strip=True) == "계약중":
        price = "계약중"
    else:
        num = item.select_one(_cls("num"))
        area = item.select_one(_cls("price"))
        if num:
            price = f"{num.get_text(strip=True)}만원"
        elif area:
            txt = area.get_text(" ", strip=True)
            price = "계약중" if ("계약중" in txt and "9,999" in txt) else txt.replace("계약중", "").strip() or txt
        else:
            m = PRICE_RE.search(item.get_text(" ", strip=True))
            price = re.sub(r"\s+", "", m.group(0)) if m else ("계약중" if "계약중" in item.get_text() else "N/A")

    # --- info line (year / mileage / fuel / location) classified by content, not position ---
    lis = item.select(_cls("info") + " li") or item.select("li")
    year_txt = mileage = "N/A"
    rest = []
    for li in lis:
        s = li.get_text(" ", strip=True)
        if not s:
            continue
        if year_txt == "N/A" and YEAR_RE.search(s):
            year_txt = s
        elif mileage == "N/A" and "km" in s.lower():
            mileage = s
        else:
            rest.append(s)
    fuel_type = rest[0] if len(rest) > 0 else "N/A"
    location = rest[1] if len(rest) > 1 else "N/A"

    return {
        "car_id": car_id, "title": title, "price": price, "year": year_txt,
        "mileage": mileage, "fuel_type": fuel_type, "location": location,
        "url": DETAIL_URL.format(car_id), "image_url": _pick_image(item),
    }


def parse_listing_html(html):
    """Returns (cars, diagnostics). Tries the named card class first, then falls back
    to grouping by every <a href='.../cars/detail/ID'> so a class rename can't zero us out."""
    soup = BeautifulSoup(html, "html.parser")
    cards = soup.select(_cls("item"))
    method = "class ItemBigImage_item__*"
    if not cards:
        method = "detail-link fallback"
        cards, seen = [], set()
        for a in soup.select('a[href*="/cars/detail/"], a[href*="carid="]'):
            box = a.find_parent("li") or a.parent
            if id(box) not in seen:
                seen.add(id(box))
                cards.append(box)

    cars, ids = [], set()
    for c in cards:
        try:
            d = _parse_item(c)
        except Exception:
            d = None
        if d and d["car_id"] not in ids:
            ids.add(d["car_id"])
            cars.append(d)

    seen_cls = sorted({k for tag in soup.find_all(class_=True) for k in tag.get("class", []) if "Item" in k})[:8]
    diag = {
        "method": method,
        "cards": len(cards),
        "parsed": len(cars),
        "title": soup.title.get_text(strip=True) if soup.title else "",
        "detail_links": len(soup.select('a[href*="/cars/detail/"]')),
        "sample_classes": seen_cls,
    }
    return cars, diag


def _parse_api_car(row):
    """Normalize one result from Encar's own mobile search API."""
    car_id = str(row.get("Id") or "").strip()
    if not car_id:
        return None

    title = " ".join(str(row.get(key) or "").strip()
                     for key in ("Manufacturer", "Model", "Badge")
                     if str(row.get(key) or "").strip()) or "N/A"

    try:
        year_code = int(float(row.get("Year") or 0))
    except (TypeError, ValueError):
        year_code = 0
    if year_code >= 100000:
        year, month = divmod(year_code, 100)
    else:
        year, month = year_code, 0
    if not (1900 <= year <= MAX_YEAR):
        try:
            year = int(str(row.get("FormYear") or "0")[:4])
        except (TypeError, ValueError):
            year = 0
        month = 0
    year_txt = f"{year}/{month:02d}식" if year and 1 <= month <= 12 else (f"{year}년" if year else "N/A")

    try:
        mileage_value = int(float(row.get("Mileage") or 0))
        mileage = f"{mileage_value:,}km" if mileage_value > 0 else "N/A"
    except (TypeError, ValueError):
        mileage = "N/A"

    try:
        price_value = float(row.get("Price") or 0)
        price = f"{price_value:,.0f}만원" if price_value > 0 else "N/A"
    except (TypeError, ValueError):
        price = "N/A"

    image_path = ""
    photos = row.get("Photos") or []
    if isinstance(photos, list):
        for photo in photos:
            if isinstance(photo, dict) and photo.get("location"):
                image_path = str(photo["location"])
                break
    image_url = ("https://ci.encar.com" + image_path) if image_path.startswith("/") else image_path

    return {
        "car_id": car_id,
        "title": title,
        "price": price,
        "year": year_txt,
        "mileage": mileage,
        "fuel_type": str(row.get("FuelType") or "N/A"),
        "location": str(row.get("OfficeCityState") or "N/A"),
        "url": DETAIL_URL.format(car_id),
        "image_url": image_url,
    }


class EncarApiPageWorker(QThread):
    """Fetch later result pages by Encar's supported offset API, not its fragile UI pager."""
    page_ready = pyqtSignal(int, object, object)

    def __init__(self, page, parent=None):
        super().__init__(parent)
        self.page = page

    def run(self):
        start = (self.page - 1) * 200
        params = urlencode({
            "count": "true",
            "q": "(And.Hidden.N._.CarType.A.)",
            "sr": f"|MobileModifiedDate|{start}|200",
            "inav": "|Metadata|Sort",
            "cursor": "",
        })
        request = Request(
            f"https://api.encar.com/search/car/list/mobile?{params}",
            headers={
                "User-Agent": "Mozilla/5.0",
                "Referer": "https://car.encar.com/",
                "Origin": "https://car.encar.com",
                "Accept": "application/json, text/plain, */*",
            },
        )
        try:
            with urlopen(request, timeout=35) as response:
                payload = json.loads(response.read().decode("utf-8"))
            self.page_ready.emit(self.page, payload, None)
        except Exception as exc:
            self.page_ready.emit(self.page, None, str(exc))


# ---------- FilterDialog ----------
class FilterDialog(QDialog):
    def __init__(self, parent=None, filter_data=None):
        super().__init__(parent)
        self.setWindowTitle(tr("Shto/Ndrysho Filtër"))
        self.resize(500, 400)
        self.setStyleSheet("""
            QDialog {background-color: #f0f0f0;}
            QGroupBox{font-weight:bold;border:2px solid #cccccc;border-radius:5px;margin:10px 0px;padding-top:10px;}
            QGroupBox::title{subcontrol-origin:margin;left:10px;padding:0 5px 0 5px;}
            QComboBox,QSpinBox,QLineEdit{padding:5px;border:1px solid #ccc;border-radius:3px;min-height:20px;}
            QPushButton{padding:8px 16px;background-color:#0078d4;color:white;border:none;border-radius:4px;font-weight:bold;}
            QPushButton:hover{background-color:#106ebe;}
        """)
        layout = QVBoxLayout(self)

        brand_group = QGroupBox(tr("Marka & Modeli"))
        brand_layout = QGridLayout(brand_group)
        self.brand_combo = QComboBox();
        self.brand_combo.addItem(tr("Çdo Markë"), "")
        for b in sorted(CAR_BRANDS): self.brand_combo.addItem(b, b)
        self.brand_combo.currentTextChanged.connect(self.on_brand_changed)
        self.model_combo = QComboBox();
        self.model_combo.addItem(tr("Çdo Model"), "")
        brand_layout.addWidget(QLabel(tr("Marka:")), 0, 0);
        brand_layout.addWidget(self.brand_combo, 0, 1)
        brand_layout.addWidget(QLabel(tr("Modeli:")), 1, 0);
        brand_layout.addWidget(self.model_combo, 1, 1)
        layout.addWidget(brand_group)

        kw_group = QGroupBox(tr("Fjalë kyçe shtesë"))
        kw_layout = QVBoxLayout(kw_group)
        self.keywords_edit = QLineEdit()
        self.keywords_edit.setPlaceholderText(tr("të ndara me presje"))
        kw_layout.addWidget(self.keywords_edit)
        layout.addWidget(kw_group)

        yr_group = QGroupBox(tr("Intervali i viteve"))
        yr_layout = QGridLayout(yr_group)
        self.min_year = QSpinBox();
        self.min_year.setRange(1990, MAX_YEAR);
        self.min_year.setValue(1990);
        self.min_year.setSpecialValueText(tr("Pa kufizim"))
        self.max_year = QSpinBox();
        self.max_year.setRange(1990, MAX_YEAR);
        self.max_year.setValue(MAX_YEAR);
        self.max_year.setSpecialValueText(tr("Pa kufizim"))
        self.min_month = QSpinBox();
        self.min_month.setRange(1, 12);
        self.min_month.setValue(1);
        self.max_month = QSpinBox();
        self.max_month.setRange(1, 12);
        self.max_month.setValue(12);
        yr_layout.addWidget(QLabel(tr("Viti min:")), 0, 0);
        yr_layout.addWidget(self.min_year, 0, 1)
        yr_layout.addWidget(QLabel(tr("Viti max:")), 0, 2);
        yr_layout.addWidget(self.max_year, 0, 3)
        yr_layout.addWidget(QLabel(tr("Muaji min:")), 1, 0);
        yr_layout.addWidget(self.min_month, 1, 1)
        yr_layout.addWidget(QLabel(tr("Muaji max:")), 1, 2);
        yr_layout.addWidget(self.max_month, 1, 3)
        layout.addWidget(yr_group)

        self.enabled_box = QCheckBox(tr("Aktivizo këtë filtër"));
        self.enabled_box.setChecked(True)
        layout.addWidget(self.enabled_box)

        if filter_data: self.load_filter_data(filter_data)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self.accept);
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def on_brand_changed(self):
        b = self.brand_combo.currentData()
        self.model_combo.clear();
        self.model_combo.addItem("Çdo Model", "")
        if b and b in CAR_MODELS:
            for m in CAR_MODELS[b]: self.model_combo.addItem(m, m)

    def load_filter_data(self, d):
        if "brand" in d:
            idx = self.brand_combo.findData(d["brand"])
            if idx >= 0:
                self.brand_combo.setCurrentIndex(idx);
                self.on_brand_changed()
        if "model" in d:
            idx = self.model_combo.findData(d["model"])
            if idx >= 0: self.model_combo.setCurrentIndex(idx)
        self.keywords_edit.setText(", ".join(d.get("keywords", [])))
        self.min_year.setValue(d.get("min_year", 1990))
        self.max_year.setValue(d.get("max_year", MAX_YEAR))
        self.min_month.setValue(d.get("min_month") or 1)
        self.max_month.setValue(d.get("max_month") or 12)
        self.enabled_box.setChecked(d.get("enabled", True))

    def get_data(self):
        kw = [k.strip() for k in self.keywords_edit.text().split(",") if k.strip()]
        return {
            "brand": self.brand_combo.currentData() or None,
            "model": self.model_combo.currentData() or None,
            "keywords": kw,
            "min_year": self.min_year.value() or None,
            "max_year": self.max_year.value() or None,
            "min_month": self.min_month.value(),
            "max_month": self.max_month.value(),
            "enabled": self.enabled_box.isChecked()
        }


# ---------- FilterWidget ----------
class FilterWidget(QWidget):
    filters_changed = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.filters = []
        self.setup_ui();
        self.refresh()

    def setup_ui(self):
        v = QVBoxLayout(self)
        v.addWidget(
            QLabel(tr("Krijoni filtra të saktë për kërkimin e makinave. Filloni duke shtuar një filtër më poshtë.")))
        self.tbl = QTableWidget();
        self.tbl.setColumnCount(6)
        self.tbl.setHorizontalHeaderLabels(["✓", tr("Marka"), tr("Modeli"), tr("Fjalë kyçe"), tr("Vitet"), tr("Muajt")])
        self.tbl.horizontalHeader().setStretchLastSection(True)
        self.tbl.setAlternatingRowColors(True);
        self.tbl.setSelectionBehavior(QTableWidget.SelectRows)
        v.addWidget(self.tbl)
        h = QHBoxLayout()
        for txt, slot in (("➕ Shto", self.add_filter), ("✏️ Ndrysho", self.edit_filter),
                          ("🗑️ Fshi", self.delete_filter), ("🧹 Pastro", self.clear_all)):
            b = QPushButton(txt);
            b.clicked.connect(slot);
            h.addWidget(b)
        h.addStretch();
        v.addLayout(h)
        self.setStyleSheet("""
            QTableWidget{background-color:white;alternate-background-color:#f8f9fa;border:1px solid #ccc;border-radius:4px;}
            QPushButton{padding:8px 16px;background:#0078d4;color:white;border:none;border-radius:4px;font-weight:bold;}
            QPushButton:hover{background:#106ebe;}
        """)

    def refresh(self):
        self.tbl.setRowCount(len(self.filters))
        for i, f in enumerate(self.filters):
            chk = QCheckBox();
            chk.setChecked(f.get("enabled", True))
            chk.stateChanged.connect(lambda s, i=i: self.toggle(i, s))
            self.tbl.setCellWidget(i, 0, chk)
            self.tbl.setItem(i, 1, QTableWidgetItem(f.get("brand", tr("Çdo"))))
            self.tbl.setItem(i, 2, QTableWidgetItem(f.get("model", tr("Çdo"))))
            self.tbl.setItem(i, 3, QTableWidgetItem(", ".join(f.get("keywords", []))))
            min_y = f.get("min_year");
            max_y = f.get("max_year")
            year_rng = f"{min_y}-{max_y}" if min_y is not None and max_y is not None else tr("Çdo")
            self.tbl.setItem(i, 4, QTableWidgetItem(year_rng))
            min_m = f.get("min_month");
            max_m = f.get("max_month")
            month_rng = f"{min_m}-{max_m}" if min_m is not None and max_m is not None else tr("Çdo")
            self.tbl.setItem(i, 5, QTableWidgetItem(month_rng))

    def toggle(self, i, state):
        self.filters[i]["enabled"] = (state == Qt.Checked);
        self.filters_changed.emit()

    def add_filter(self):
        dlg = FilterDialog(self)
        if dlg.exec_():
            self.filters.append(dlg.get_data());
            self.refresh();
            self.filters_changed.emit()

    def edit_filter(self):
        r = self.tbl.currentRow()
        if r < 0: QMessageBox.information(self, tr("Asnjë përzgjedhje"),
                                          tr("Ju lutem zgjidhni një filtër për ta ndryshuar.")); return
        dlg = FilterDialog(self, self.filters[r])
        if dlg.exec_():
            self.filters[r] = dlg.get_data();
            self.refresh();
            self.filters_changed.emit()

    def delete_filter(self):
        r = self.tbl.currentRow()
        if r < 0: QMessageBox.information(self, tr("Asnjë përzgjedhje"),
                                          tr("Ju lutem zgjidhni një filtër për ta fshirë.")); return
        if QMessageBox.question(self, tr("Konfirmo"), tr("Fshi filtrin?")) == QMessageBox.Yes:
            del self.filters[r];
            self.refresh();
            self.filters_changed.emit()

    def clear_all(self):
        if not self.filters: return
        if QMessageBox.question(self, tr("Konfirmo"), tr("Pastro të gjitha?")) == QMessageBox.Yes:
            self.filters.clear();
            self.refresh();
            self.filters_changed.emit()


# ---------- MainWindow ----------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.conn = init_db()
        self.current_page = 1
        self.max_pages = MAX_PAGES
        self.raw_cars = []
        self.api_worker = None
        self.setup_ui()
        self.setStyleSheet("""
            QMainWindow,QWidget{background:#f8f9fa;font-family:'Segoe UI',Arial;}
            QGroupBox{font-weight:bold;border:2px solid #dee2e6;border-radius:8px;margin:10px 0px;padding-top:15px;background:white;}
            QPushButton{padding:10px 20px;background:#0078d4;color:white;border:none;border-radius:6px;font-weight:600;}
            QPushButton:hover{background:#106ebe;} QPushButton:disabled{background:#6c757d;}
            QProgressBar{border:none;background:#e9ecef;border-radius:10px;text-align:center;font-weight:bold;}
            QProgressBar::chunk{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #28a745,stop:1 #20c997);border-radius:10px;}
        """)

    def setup_ui(self):
        self.setWindowTitle(tr("Encar Scraper Pro – Filtro PËRPARA se të ruash"))
        self.resize(1400, 900)
        main = QHBoxLayout(self)
        split = QSplitter(Qt.Horizontal)

        left = QWidget()
        lv = QVBoxLayout(left)

        cg = QGroupBox(tr("🎛️ Kontrollet e shkarkimit"))
        cv = QVBoxLayout(cg)
        h = QHBoxLayout();
        h.addWidget(QLabel(tr("Faqet:")))
        self.pages_spin = QSpinBox();
        self.pages_spin.setRange(1, 50);
        self.pages_spin.setValue(MAX_PAGES)
        h.addWidget(self.pages_spin);
        self.flag_sq = QPushButton("Shqip")
        self.flag_sq.setIcon(flag_icon("al"))
        self.flag_sq.setIconSize(QSize(42, 28))
        self.flag_en = QPushButton("English")
        self.flag_en.setIcon(flag_icon("gb"))
        self.flag_en.setIconSize(QSize(42, 28))
        self.flag_sq.setCheckable(True)
        self.flag_en.setCheckable(True)
        self.flag_sq.setChecked(True)
        self.flag_sq.clicked.connect(lambda: self.set_language("sq"))
        self.flag_en.clicked.connect(lambda: self.set_language("en"))
        h.addWidget(self.flag_sq)
        h.addWidget(self.flag_en)
        cv.addLayout(h)
        self.start_btn = QPushButton(tr("🚀 Fillo shkarkimin"));
        self.start_btn.clicked.connect(self.start_scraping)
        cv.addWidget(self.start_btn)
        self.progress = QProgressBar();
        cv.addWidget(self.progress)
        self.stats = QLabel(tr("Gati – konfiguro filtrat fillimisht"));
        self.stats.setStyleSheet("color:#28a745;font-weight:bold;padding:5px")
        cv.addWidget(self.stats)
        lv.addWidget(cg)

        fg = QGroupBox(tr("🔍 Filtrat e mençur"))
        fv = QVBoxLayout(fg)
        self.filter_widget = FilterWidget()
        fv.addWidget(self.filter_widget)
        lv.addWidget(fg)
        split.addWidget(left)

        self.tabs = QTabWidget()
        browser_tab = QWidget();
        bv = QVBoxLayout(browser_tab)
        self.web = QWebEngineView();
        self.web.load(QUrl(build_list_url(1)))
        self.web.loadFinished.connect(self._on_page_loaded)
        self.web.page().renderProcessTerminated.connect(self._on_renderer_terminated)
        bv.addWidget(self.web);
        self.tabs.addTab(browser_tab, tr("🌐 Shfletuesi"))

        results_tab = QWidget();
        rv = QVBoxLayout(results_tab)
        self.results = QListWidget();
        self.results.setAlternatingRowColors(True);
        rv.addWidget(self.results)
        rv.addWidget(QLabel(tr("📝 Ditari i shkarkimit:")));
        self.log = QTextEdit();
        self.log.setReadOnly(True);
        self.log.setMaximumHeight(120)
        self.log.setStyleSheet(
            "background:#2d3748;color:#e2e8f0;font-family:'Courier New';padding:8px;border-radius:4px")
        rv.addWidget(self.log)
        self.tabs.addTab(results_tab, tr("📊 Rezultatet"))
        split.addWidget(self.tabs)
        split.setSizes([400, 1000])
        main.addWidget(split)

    def append_log(self, message):
        self.log.append(translate_log(str(message), LANGUAGE))

    @staticmethod
    def _decode_js_result(result):
        """Normalize Qt WebEngine callback results across QVariant/QJSValue versions."""
        if hasattr(result, "toVariant"):
            result = result.toVariant()
        if isinstance(result, bytes):
            result = result.decode("utf-8", "replace")
        if isinstance(result, str):
            try:
                result = json.loads(result)
            except (TypeError, ValueError):
                return {}
        return result if isinstance(result, dict) else {}

    def _on_renderer_terminated(self, status, exit_code):
        self.append_log(f"⚠️ Web engine stopped (status={status}, code={exit_code}).")
        if getattr(self, "scraping", False):
            self.scraping = False
            self.start_btn.setEnabled(True)
            self.start_btn.setText(tr("🚀 Fillo shkarkimin"))

    def set_language(self, language):
        global LANGUAGE
        LANGUAGE = language
        self.flag_sq.setChecked(language == "sq")
        self.flag_en.setChecked(language == "en")
        # Update the existing widgets in place so filters, results, and progress are preserved.
        reverse = {english: albanian for albanian, english in TRANSLATIONS.items()}
        for widget in self.findChildren(QWidget):
            if isinstance(widget, QGroupBox):
                value = widget.title()
                widget.setTitle(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QLabel):
                value = widget.text()
                widget.setText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QPushButton) and widget not in (self.flag_sq, self.flag_en):
                value = widget.text()
                widget.setText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QCheckBox):
                value = widget.text()
                widget.setText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QLineEdit):
                value = widget.placeholderText()
                widget.setPlaceholderText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QSpinBox):
                value = widget.specialValueText()
                widget.setSpecialValueText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QTabWidget):
                for index in range(widget.count()):
                    value = widget.tabText(index)
                    widget.setTabText(index, TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
            elif isinstance(widget, QTableWidget):
                for col in range(widget.columnCount()):
                    item = widget.horizontalHeaderItem(col)
                    if item:
                        value = item.text()
                        item.setText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
                for row in range(widget.rowCount()):
                    for col in range(1, widget.columnCount()):
                        item = widget.item(row, col)
                        if item:
                            value = item.text()
                            item.setText(TRANSLATIONS.get(value, value) if language == "en" else reverse.get(value, value))
        title = "Encar Scraper Pro – Filter before saving" if language == "en" else "Encar Scraper Pro – Filtro PËRPARA se të ruash"
        self.setWindowTitle(title)
        self.filter_widget.refresh()
        self.log.setPlainText(translate_log(self.log.toPlainText(), language))
        if self.start_btn.isEnabled():
            self.start_btn.setText(tr("🚀 Fillo shkarkimin"))

    def start_scraping(self):
        self.raw_cars.clear();
        self.results.clear();
        self.log.clear()
        self.max_pages = self.pages_spin.value()
        self.current_page = 1
        self.scraping = True
        self.progress.setValue(0)
        self.start_btn.setEnabled(False)
        self.start_btn.setText(tr("⏳ Duke shkarkuar..."))
        self.stats.setText(tr("Duke shkarkuar …"))
        self.previous_page_ids = set()
        self.page_navigation_retries = 0
        self.append_log("🚀 Filloi shkarkimi i ri")

        # TEST YEAR EXTRACTION në fillim
        test_year_extraction()

        self.load_next()

    def load_next(self):
        if self.current_page > self.max_pages:
            self.finish_scraping()
            return
        if self.current_page == 1:
            url = build_list_url(1)
            self.append_log(f"🌐 Duke ngarkuar faqen 1: {url}")
            self.poll_ticks = 0
            self.waiting_for_page = True
            self.web.load(QUrl(url))
            return

        offset = (self.current_page - 1) * 200
        self.append_log(
            f"🌐 Duke kërkuar faqen {self.current_page} nga Encar API (offset={offset})."
        )
        self.api_worker = EncarApiPageWorker(self.current_page, self)
        self.api_worker.page_ready.connect(self._on_api_page_ready)
        self.api_worker.start()

    def _on_api_page_ready(self, page, payload, error):
        if not self.scraping or page != self.current_page:
            return
        if error:
            self._retry_page_navigation(f"Encar API request failed: {error}")
            return

        rows = payload.get("SearchResults", []) if isinstance(payload, dict) else []
        cars = []
        for row in rows:
            try:
                car = _parse_api_car(row)
                if car:
                    cars.append(car)
            except Exception as exc:
                self.append_log(f"⚠️ Gabim gjatë leximit të rezultatit nga API: {exc}")

        if not cars and int(payload.get("Count", 0) or 0) > (page - 1) * 200:
            self._retry_page_navigation(f"Encar API returned no cars for offset {(page - 1) * 200}")
            return

        start = (page - 1) * 200
        total = payload.get("Count", "?")
        self.append_log(
            f"🔎 Encar API: {len(cars)} makina (offset {start}, total {total})"
        )
        self.process_cars(cars)

    def _retry_page_navigation(self, reason):
        """Retry the site's client-side pagination; never report a partial run as complete."""
        if self.page_navigation_retries < 3:
            self.page_navigation_retries += 1
            self.append_log(
                f"⚠️ Navigimi te faqja {self.current_page} dështoi "
                f"({self.page_navigation_retries}/3): {reason}. Po riprovohet..."
            )
            self.poll_ticks = 0
            QTimer.singleShot(1800, self.load_next)
            return

        self.scraping = False
        self.start_btn.setEnabled(True)
        self.start_btn.setText(tr("🚀 Fillo shkarkimin"))
        self.stats.setText(
            f"⚠️ Ndaloi te faqja {self.current_page}; faqet e mëparshme janë ruajtur."
            if LANGUAGE == "sq" else
            f"⚠️ Stopped at page {self.current_page}; earlier pages were saved."
        )
        self.append_log(
            f"❌ Nuk u arrit faqja {self.current_page} pas 3 përpjekjeve. "
            "Aplikacioni mbetet i hapur; kontrolloni shfletuesin dhe nisni përsëri."
            if LANGUAGE == "sq" else
            f"❌ Could not reach page {self.current_page} after 3 attempts. "
            "The app is still open; check the browser and start another run."
        )

    def _on_page_loaded(self, ok):
        """Do not inspect HTML from the previous page while a new URL is still loading."""
        if not getattr(self, "scraping", False) or not getattr(self, "waiting_for_page", False):
            return
        self.waiting_for_page = False
        if not ok:
            self.append_log(f"⚠️ Faqja {self.current_page} nuk u ngarkua; po provohet leximi i saj.")
        self.poll_ticks = 0
        QTimer.singleShot(900, self.poll_page)

    def poll_page(self):
        """Wait until listing links actually exist (the page is a JS app), instead of a blind sleep."""
        js = """(() => {
            const cards = [...document.querySelectorAll('[class*="ItemBigImage_item__"]')];
            const ids = cards.map(card => {
                const node = card.matches('[data-enlog-dt-param]') ? card : card.querySelector('[data-enlog-dt-param]');
                if (node) { try { return String(JSON.parse(node.getAttribute('data-enlog-dt-param')).carId || ''); } catch(e) {} }
                const a = card.matches('a[href*="/cars/detail/"]') ? card : card.querySelector('a[href*="/cars/detail/"]');
                const m = a && a.href.match(/\/cars\/detail\/(\d+)/);
                return m ? m[1] : '';
            }).filter(Boolean);
            return JSON.stringify({count: cards.length || document.querySelectorAll('a[href*="/cars/detail/"]').length,
                ids, url:location.href, text:(document.body && document.body.innerText || '').slice(-1200)});
        })()"""
        self.web.page().runJavaScript(js, self._on_poll)

    def _on_poll(self, result):
        self.poll_ticks += 1
        result = self._decode_js_result(result)
        count = result.get("count", 0)
        ids = set(result.get("ids", []))
        body_text = result.get("text", "")
        url = result.get("url", "")
        page_match = re.search(r"[?&]page=(\d+)", url)
        url_page = int(page_match.group(1)) if page_match else 1
        no_results = any(marker in body_text for marker in ("검색 결과가 없습니다", "검색된 차량이 없습니다", "조회된 차량이 없습니다"))
        page_changed = self.current_page == 1 or bool(ids - self.previous_page_ids)
        if self.current_page > 1 and url_page != self.current_page and self.poll_ticks < 40:
            QTimer.singleShot(700, self.poll_page)
            return
        if self.current_page > 1 and url_page != self.current_page:
            self._retry_page_navigation(f"page URL did not change (current URL: {url})")
            return
        # Encar changes the route before its async results arrive. Allow that fetch to settle;
        # if cards are present after the wait, process them even if some IDs overlap.
        if self.current_page > 1 and not page_changed and count > 0 and not no_results and self.poll_ticks < 10:
            QTimer.singleShot(700, self.poll_page)
            return
        if self.current_page > 1 and count == 0 and not no_results and self.poll_ticks < 40:
            QTimer.singleShot(700, self.poll_page)
            return
        if self.current_page > 1 and count == 0 and not no_results:
            self._retry_page_navigation(f"no listing cards returned (URL: {url})")
            return
        if count > 0 or no_results or self.poll_ticks >= 30:
            if count == 0 and not no_results:
                self.append_log(f"⚠️ Faqja {self.current_page}: të dhënat nuk u shfaqën pas 30 sekondash.")
            self._scroll_then_scrape(0)
        else:
            QTimer.singleShot(700, self.poll_page)

    def _scroll_then_scrape(self, step):
        """Scroll down in steps so lazy-loaded thumbnails and cards get rendered."""
        if step >= 4:
            self.scrape_current()
            return
        self.web.page().runJavaScript(
            f"window.scrollTo(0, document.body.scrollHeight * {(step + 1) / 4});")
        QTimer.singleShot(400, lambda: self._scroll_then_scrape(step + 1))

    def scrape_current(self):
        self.web.page().toHtml(self.process_html)

    def process_html(self, html):
        cars, diag = parse_listing_html(html or "")
        self.append_log(f"🔎 Parser: {diag['parsed']} makina ({diag['method']}), "
                        f"links={diag['detail_links']}, title='{diag['title'][:40]}'")
        self.previous_page_ids = {str(car["car_id"]) for car in cars}
        if not cars:
            dump = Path(f"debug_page_{self.current_page}.html")
            try:
                dump.write_text(html or "", encoding="utf-8")
            except Exception:
                pass
            self.append_log(f"⚠️ 0 makina. HTML u ruajt te {dump} | klasat: {diag['sample_classes']}")

        self.process_cars(cars)

    def process_cars(self, cars):
        """Apply the same filters and storage path to HTML and API-parsed cars."""
        filters = self.filter_widget.filters
        total = 0
        brand_stats = {}

        for car_data in cars:
            try:
                car_data["first_seen"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                title, year_txt = car_data["title"], car_data["year"]
                price, mileage = car_data["price"], car_data["mileage"]
                fuel_type, location = car_data["fuel_type"], car_data["location"]
                img_url = car_data["image_url"]

                title_lower = title.lower()
                for brand, keywords in CAR_BRANDS.items():
                    if any(k.lower() in title_lower for k in keywords):
                        brand_stats[brand] = brand_stats.get(brand, 0) + 1
                        break

                passes_filter = matches_filters_strict(title, year_txt, filters) if filters else True

                if filters and self.current_page <= 2:
                    year, _m = extract_year_month(year_txt)
                    self.append_log(f"🔍 {title[:30]}... | Year: {year} | Passes: {passes_filter}")

                if not passes_filter:
                    continue

                self.conn.execute(
                    "INSERT OR IGNORE INTO seen_cars VALUES (?,?,?,?,?,?,?,?)",
                    (car_data["car_id"], car_data["url"], title, price, mileage,
                     year_txt, car_data["first_seen"], img_url))
                self.conn.commit()
                save_excel(car_data)
                self.raw_cars.append(car_data)

                display_price = price if price != "N/A" else "Çmimi i padisponueshëm"
                img_status = "🖼️" if img_url else "❌"
                extracted_year, _ = extract_year_month(year_txt)
                year_display = (f"📅 {year_txt.replace('식', '')} ({extracted_year})"
                                if extracted_year else f"📅 {year_txt.replace('식', '')}")
                item = QListWidgetItem(
                    f"✅ {title} | 💰 {display_price} | {year_display} | "
                    f"🏃 {mileage} | ⛽ {fuel_type} | 📍 {location} | {img_status}")
                item.setBackground(QColor(220, 255, 220))
                self.results.addItem(item)
                total += 1
            except Exception as e:
                self.append_log(f"⚠️ Gabim gjatë procesimit: {str(e)}")

        self.append_log(f"📋 Faqja {self.current_page}: {total} makina u ruajtën pas filtrimit")
        if brand_stats:
            self.append_log("📊 Markat në faqe: " + " | ".join(f"{b}: {c}" for b, c in sorted(brand_stats.items())))

        self.current_page += 1
        self.page_navigation_retries = 0
        self.progress.setValue(int((self.current_page - 1) * 100 / self.max_pages))
        QTimer.singleShot(800, self.load_next)

    def finish_scraping(self):
        self.scraping = False
        self.progress.setValue(100)
        self.start_btn.setEnabled(True);
        self.start_btn.setText(tr("🚀 Fillo shkarkimin"))
        car_word = ("makinë" if len(self.raw_cars) == 1 else "makina") if LANGUAGE == "sq" else ("car" if len(self.raw_cars) == 1 else "cars")
        self.stats.setText((f"✅ Përfunduar – {len(self.raw_cars)} {car_word} u mbajt pas filtrimit" if LANGUAGE == "sq"
                            else f"✅ Finished – {len(self.raw_cars)} {car_word} kept after filtering"))

        # STATISTIKA FINALE
        final_brand_stats = {}
        for car in self.raw_cars:
            title_lower = car["title"].lower()
            for brand, keywords in CAR_BRANDS.items():
                if any(k.lower() in title_lower for k in keywords):
                    final_brand_stats[brand] = final_brand_stats.get(brand, 0) + 1
                    break

        stats_msg = ((f"Shkarkim i përfunduar.\n{len(self.raw_cars)} makina u ruajtën.\n\nShpërndarja e markave:\n") if LANGUAGE == "sq"
                     else f"Scraping complete.\n{len(self.raw_cars)} cars saved.\n\nBrand breakdown:\n")
        for brand, count in sorted(final_brand_stats.items()):
            unit = ("makinë" if count == 1 else "makina") if LANGUAGE == "sq" else ("car" if count == 1 else "cars")
            stats_msg += f"• {brand}: {count} {unit}\n"

        QMessageBox.information(self, tr("Përfunduar"), stats_msg)

    def closeEvent(self, event):
        if self.conn: self.conn.close()
        event.accept()


# ---------- Run ----------
def main():
    app = QApplication(sys.argv)
    def report_unhandled_exception(exc_type, exc_value, exc_tb):
        try:
            crash_file = APP_DIR / "autokorea_crash.log"
            with crash_file.open("a", encoding="utf-8") as log_file:
                log_file.write("\n--- Unhandled application exception ---\n")
                traceback.print_exception(exc_type, exc_value, exc_tb, file=log_file)
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc_value, exc_tb)
    sys.excepthook = report_unhandled_exception
    app.setApplicationName("Encar Scraper Pro")
    app.setApplicationVersion("4.0")
    app.setOrganizationName("CarScraper")
    w = MainWindow();
    w.show()
    return app.exec_()


if __name__ == "__main__":
    sys.exit(main())
