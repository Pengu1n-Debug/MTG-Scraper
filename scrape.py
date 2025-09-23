import httpx
import threading
import time
import webbrowser
import os
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = r"C:\Users\kaden\PlaywrightBrowsers"
import datetime
import sys
from urllib.parse import urlencode, quote, quote_plus
import cloudscraper
from bs4 import BeautifulSoup
from parsel import Selector
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import requests
import re
import json
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook
from playwright.sync_api import sync_playwright
from tkinterdnd2 import TkinterDnD, DND_FILES
import mtg_parser

SORTING_MAP = {'best_match': 12}
session = httpx.Client(headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br'
}, http2=True, follow_redirects=True, timeout=30.0)

def scrape_moonmtg(card_name):
    BASE_URL = 'https://moonmtg.com/products/'
    import re
    import requests
    from bs4 import BeautifulSoup

    def name_to_handle(card_name):
        handle = card_name.lower()
        handle = re.sub(r"[’'\":,?!()]", '', handle)
        handle = re.sub(r'[^a-z0-9\s-]', '', handle)
        handle = re.sub(r'\s+', '-', handle)
        handle = handle.strip('-')
        return handle

    def normalize_name(name):
        name = name.lower()
        name = re.sub(r"[’'\":,?!()]", '', name)
        name = re.sub(r'[^a-z0-9\s]', '', name)
        return name.strip()

    def fetch_product_json(handle):
        url = f'{BASE_URL}{handle}.json'
        try:
            response = requests.get(url)
            if response.status_code == 200:
                return response.json()
        except:
            pass
        return None

    def fetch_variant_stock(handle, variant_id):
        variant_url = f'{BASE_URL}{handle}?variant={variant_id}'
        try:
            response = requests.get(variant_url)
            if response.status_code != 200:
                return 'Unknown'
            soup = BeautifulSoup(response.text, 'html.parser')
            inventory_element = soup.find('p', class_='product__inventory')
            return inventory_element.get_text(strip=True) if inventory_element else 'Stock info not found'
        except:
            return 'Unknown'

    handle = name_to_handle(card_name)
    product_json = fetch_product_json(handle)
    success = False
    if product_json:
        product = product_json['product']
        if normalize_name(card_name) in normalize_name(product['title']):
            success = True
    if not success:
        handle += '-1'
        product_json = fetch_product_json(handle)
        if not product_json:
            return (0.0, 'Not found', '')
        product = product_json['product']
        if normalize_name(card_name) not in normalize_name(product['title']):
            return (0.0, 'Not found', '')
    in_stock_variants = []
    for variant in product['variants']:
        price = float(variant['price'])
        variant_id = variant['id']
        stock_status = fetch_variant_stock(handle, variant_id)
        if stock_status not in ['Out of stock', 'Stock info not found', 'Unknown']:
            in_stock_variants.append((variant['title'], price, variant_id))
    if not in_stock_variants:
        return (0.0, 'Out of stock', '')
    title, price, variant_id = sorted(in_stock_variants, key=lambda x: x[1])[0]
    return (price, title, f'{BASE_URL}{handle}?variant={variant_id}')

def fetch_mtgmate_price(card_name: str):
    """
    Scrape MTGMate search results for a given card.
    Returns (cheapest_price, title, url) like other scrapers.
    """
    url = f"https://www.mtgmate.com.au/cards/search?q={card_name.replace(' ', '+')}"
    scraper = cloudscraper.create_scraper()

    try:
        r = scraper.get(url, timeout=20)
        r.raise_for_status()
    except Exception as e:
        print(f"[MTGMate] Request failed: {e}")
        return (0.0, "Error", "")

    soup = BeautifulSoup(r.text, "html.parser")
    container = soup.find("div", {"data-react-class": "FilterableTable"})
    if not container:
        print("[MTGMate] Could not find FilterableTable div.")
        return (0.0, "Not found", "")

    raw_props = container.get("data-react-props")
    if not raw_props:
        print("[MTGMate] No data-react-props found.")
        return (0.0, "Not found", "")

    try:
        data = json.loads(raw_props)
    except Exception as e:
        print(f"[MTGMate] JSON parsing error: {e}")
        return (0.0, "Error", "")

    uuid_map = data.get("uuid", {})
    results = []

    for card in data.get("cards", []):
        card_id = card.get("uuid")
        details = uuid_map.get(card_id, {})
        if not details:
            continue

        try:
            price = int(details.get("price", 0)) / 100
        except (TypeError, ValueError):
            price = 0.0

        qty = details.get("quantity", 0)
        if price > 0 and qty > 0:
            results.append((
                price,
                f"{details.get('name')} ({details.get('set_name')}, {details.get('finish')})",
                f"https://www.mtgmate.com.au{details.get('link_path', '')}"
            ))

    if not results:
        return (0.0, "Out of stock", "")

    cheapest = min(results, key=lambda x: x[0])
    print(f"[MTGMate] Cheapest: {cheapest}")
    return cheapest

def scrape_gg(card_name, base_url):
    def normalize(text):
        text = text.lower()
        text = re.sub(r"[’'\":,?!()\[\]]", "", text)
        text = re.sub(r"[^a-z0-9\s\-]", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    target = normalize(card_name)
    target_first = target.split()[0] if target else ""
    query = quote_plus(f"{card_name} product_type:\"mtg\"")
    url = f"{base_url}/search?q={query}"
    headers = {"User-Agent": "Mozilla/5.0"}

    print(f"\n[GG] Searching for: {card_name}")
    print(f"[GG] Visiting URL: {url}")

    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        results = []
        items = soup.select("div.addNow.single")
        print(f"[GG] Found {len(items)} product blocks")

        for idx, div in enumerate(items, 1):
            onclick = div.get("onclick", "")
            match = re.search(r"addToCart\([^,]+,'([^']+)'", onclick)
            title = match.group(1).strip() if match else "N/A"

            price_tag = div.find("p")
            price_text = price_tag.get_text(strip=True) if price_tag else "N/A"
            price_match = re.search(r"\$([\d.,]+)", price_text)
            price = float(price_match.group(1).replace(",", "")) if price_match else 0.0

            title_norm = normalize(title)
            title_first = title_norm.split()[0] if title_norm else ""

            print(f"[GG] #{idx} Title: {title} | Price: {price_text} | Parsed: {price}")
            if title_first != target_first:
                print(f"[GG] Skipping: '{title_first}' != '{target_first}'")
                continue

            results.append((price, title, url))

        if not results:
            print("[GG] No valid GoodGames results found")
            return 0.0, "Out of stock", ""

        cheapest = min(results, key=lambda x: x[0])
        print(f"[GG] Cheapest GoodGames: {cheapest}")
        return cheapest

    except Exception as e:
        print(f"[GG] {e}")
        return 0.0, "Error", ""

def clean_name(title: str) -> str:
    """Base card name: part before first '(' or '[' – lowercase, trimmed."""
    base = re.split(r'[\(\[]', title, 1)[0]
    return base.strip().lower()

import re
import requests
from bs4 import BeautifulSoup

def scrape_cardhub(card_name):
    def normalize(text):
        text = text.lower()
        text = re.sub(r"[^a-z0-9\s-]", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    target = normalize(card_name)

    url = f"https://thecardhubaustralia.com.au/search?type=product&options%5Bprefix%5D=last&q={card_name.replace(' ', '+')}"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        results = []
        items = soup.select("div.h4.grid-view-item__title")
        print(f"[CardHub] Searching for: {card_name}")
        print(f"[CardHub] Found {len(items)} product titles")

        for idx, title_div in enumerate(items, 1):
            title = title_div.get_text(strip=True)
            price_tag = title_div.find_next("span", class_="product-price__price")
            if not price_tag:
                print(f"[CardHub Error] Skipping #{idx}, no price tag")
                continue

            price_match = re.search(r"\$([\d.,]+)", price_tag.get_text())
            if not price_match:
                continue
            price = float(price_match.group(1).replace(",", ""))

            title_norm = normalize(title.split("(")[0].split("[")[0])

            if title_norm != target:
                print(f"[CardHub] Skipping #{idx}, title mismatch: '{title_norm}' != '{target}'")
                continue

            link_tag = title_div.find_parent("a")
            link = link_tag["href"] if link_tag else ""
            if link and not link.startswith("http"):
                link = "https://thecardhubaustralia.com.au" + link

            try:
                product_resp = requests.get(link, headers=headers, timeout=15)
                product_resp.raise_for_status()
                product_soup = BeautifulSoup(product_resp.text, "html.parser")

                if product_soup.select_one(".product-info.product-soldout"):
                    print(f"[CardHub] Skipping #{idx}, product-soldout container found: {title}")
                    continue

                add_to_cart_btn = product_soup.select_one(".product-form__item--submit button")
                if add_to_cart_btn:
                    btn_text = add_to_cart_btn.get_text(strip=True).lower()
                    is_disabled = add_to_cart_btn.has_attr("disabled")
                    if "sold out" in btn_text or is_disabled:
                        print(f"[CardHub] Skipping #{idx}, sold out via button: {title}")
                        continue
                else:
                    print(f"[CardHub Warning] Could not find add-to-cart button on {link}")
                    continue

            except Exception as e:
                print(f"[CardHub Error] Failed stock check for {link}: {e}")
                continue

            results.append((price, title, link))

        if not results:
            print("[CardHub] No valid CardHub results found")
            return 0.0, "Out of stock", ""

        cheapest = min(results, key=lambda x: x[0])
        print(f"[CardHub] Cheapest CardHub: {cheapest}")
        return cheapest

    except Exception as e:
        print(f"[CardHub] {e}")
        return 0.0, "Error", ""

def scrape_ggadelaide(card_name: str):
    return scrape_gg(card_name, base_url="https://ggadelaide.com.au")


def scrape_ggmodbury(card_name: str):
    return scrape_gg(card_name, base_url="https://ggmodbury.com.au")

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def scrape_ggaustralia(card_name: str):
    def normalize(name: str) -> str:
        name = re.split(r'[\(\[]', name)[0]
        name = name.lower()
        name = re.sub(r"[’'\":,?!()\[\]]", "", name)
        name = re.sub(r"[^a-z0-9\s\-]", "", name)
        name = re.sub(r"\s+", " ", name)
        return name.strip()

    target_normalized = normalize(card_name)

    url = (
        f"https://tcg.goodgames.com.au/search?q={card_name.replace(' ', '+')}"
        f"&s=-isActive,new_discounted_price,-_rank&f_Availability=Exclude+Out+Of+Stock"
    )

    print(f"\n[GGAustralia] Searching for: {card_name}")
    print(f"[GGAustralia] Visiting URL: {url}")

    try:
        options = Options()
        options.add_argument("--headless=new")   
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")

        driver = webdriver.Chrome(options=options)
        driver.get(url)

        try:
            WebDriverWait(driver, 12).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".st-product"))
            )
        except Exception:
            print("[GGAustralia] Product containers did not appear in time.")

        soup = BeautifulSoup(driver.page_source, "html.parser")
        product_containers = soup.select(".st-product")

        if not product_containers:
            print("[GGAustralia] No product containers found. Page source preview:")
            print(driver.page_source[:2000])  

        driver.quit()

        print(f"[GGAustralia] Found {len(product_containers)} product containers")

        results = []
        for i, prod in enumerate(product_containers, 1):
            print(f"\n[GGAustralia] --- Product #{i} ---")

            title_tag = prod.select_one(".product-title span")
            title = title_tag.get_text(strip=True) if title_tag else "N/A"
            normalized_title = normalize(title)

            if normalized_title != target_normalized:
                print(f"[GGAustralia] Skipping: '{title}' does not match '{card_name}'")
                continue

            price_tag = (
                prod.select_one(".price.no_sale")
                or prod.select_one(".discounted_price")
                or prod.select_one(".price")
            )
            price_str = price_tag.get_text(strip=True) if price_tag else None

            link_tag = prod.select_one(".product-title a")
            link = (
                link_tag["href"]
                if link_tag and "href" in link_tag.attrs
                else "https://tcg.goodgames.com.au"
            )

            print(f"[GGAustralia] Title: {title}")
            print(f"[GGAustralia] Price: {price_str}")
            print(f"[GGAustralia] Link: {link}")

            if not (title and price_str and link):
                print("[GGAustralia] Skipping: Missing required info.")
                continue

            match = re.search(r"\$([\d,]+\.\d{2})", price_str)
            if not match:
                print(f"[GGAustralia] Couldn't parse numeric price from: {price_str}")
                continue

            price = float(match.group(1).replace(",", ""))
            if not link.startswith("http"):
                link = "https://tcg.goodgames.com.au" + link

            results.append((price, title, link))

        if not results:
            print("[GGAustralia] No valid matching products with parsable price.")
            return 0.0, "No valid match", ""

        cheapest = min(results, key=lambda x: x[0])
        return cheapest

    except Exception as e:
        print(f"[GGAustralia scrape error]: {e}")
        return 0.0, "Error", ""
    
def scrape_jenes(card_name: str):
    import requests, re
    from bs4 import BeautifulSoup

    url = f"https://jenesmtg.com.au/search?q={card_name}&options%5Bprefix%5D=last"
    headers = {"User-Agent": "Mozilla/5.0"}

    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        target = card_name.strip().lower()
        results = []

        for card in soup.select("div.card-wrapper"):
            if card.select_one("span.badge") and "Sold out" in card.get_text():
                continue

            name_tag = card.select_one("a.full-unstyled-link")
            if not name_tag:
                continue

            full_name = name_tag.get_text(strip=True)
            product_name = full_name.split("|")[0].strip().lower()

            if product_name != target:
                continue

            link = name_tag.get("href", "")
            if link and not link.startswith("http"):
                link = "https://jenesmtg.com.au" + link

            found_prices = set()
            for price_tag in card.select("span.price-item"):
                text = price_tag.get_text(strip=True)
                match = re.search(r"\$([0-9]+\.[0-9]{2})", text)
                if match:
                    found_prices.add(float(match.group(1)))

            if found_prices:
                cheapest = min(found_prices)
                results.append((cheapest, full_name, link))

        if not results:
            return (0.0, "Out of stock", "")

        cheapest = min(results, key=lambda x: x[0])
        return cheapest

    except Exception as e:
        print(f"[Jene's scrape error]: {e}")
        return (0.0, "Error", "")

def parse_decklist_from_input(text):
    cards = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        match = re.match(r'(\d+x?\s*)?(.*)', line, re.IGNORECASE)
        if match:
            card_name = match.group(2).strip()
            if card_name:
                cards.append(card_name)
    return cards

CACHE_FILE = os.path.join(os.path.dirname(__file__), "deck_cache.json")

def load_deck_cache():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_deck_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)

SCRAPER_CONFIG = {
    "MoonMTG": {"enabled": True, "func": scrape_moonmtg}, 
    "MTGMate": {"enabled": True, "func": fetch_mtgmate_price},
    "CardHub": {"enabled": True, "func": scrape_cardhub},
    "JenesMTG": {"enabled": True, "func": scrape_jenes},
    "GGAustralia": {"enabled": True, "func": scrape_ggadelaide},
    "GGModbury": {"enabled": True, "func": scrape_ggmodbury},
    "GGAdelaide": {"enabled": True, "func": scrape_ggadelaide}, 
}

SOURCE_TO_COLUMN = {
    "MoonMTG": "Moon",
    "MTGMate": "MTGMate",
    "CardHub": "CardHub",
    "JenesMTG": "Jenes",
    "GGAustralia": "GGTCG",
    "GGModbury": "GGModbury",
    "GGAdelaide": "GoodGames",
}

from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, webbrowser, os, datetime, time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook

from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, webbrowser, os, datetime, time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook

from tkinterdnd2 import TkinterDnD, DND_FILES
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, webbrowser, os, datetime, time
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook

class MTGScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("MTG Price Checker")
        self.card_urls = {}
        self.stop_flag = False
        headers = {"user-agent": "my-mtg-scraper/1.0 (contact: kadenschaedel@gmail.com)"}
        self.http_client = httpx.Client(headers=headers)

        toolbar = tk.Frame(root, bd=1, relief="raised")
        toolbar.pack(side="top", fill="x")

        self.quick_menu_button = tk.Menubutton(toolbar, text="Open Sources", relief="raised")
        self.quick_menu = tk.Menu(self.quick_menu_button, tearoff=0)
        for source in SCRAPER_CONFIG:
            self.quick_menu.add_command(
                label=f"From {source}",
                command=lambda s=source: self.open_cheapest_from_source(s)
            )
        self.quick_menu.add_command(label="All from All Sources", command=self.open_all_cheapest_by_source)
        self.quick_menu_button.config(menu=self.quick_menu)
        self.quick_menu_button.pack(side="left", padx=5, pady=2)

        self.toggles_button = tk.Menubutton(toolbar, text="Toggles", relief="raised")
        self.toggles_menu = tk.Menu(self.toggles_button, tearoff=0)
        self.source_vars = {}
        for source in SCRAPER_CONFIG:
            var = tk.BooleanVar(value=SCRAPER_CONFIG[source]['enabled'])
            self.source_vars[source] = var
            self.toggles_menu.add_checkbutton(label=source, variable=var, command=self.recalculate_cheapest_prices)

        self.include_sideboard = tk.BooleanVar(value=False)
        self.include_maybeboard = tk.BooleanVar(value=False)
        self.toggles_menu.add_separator()
        self.toggles_menu.add_checkbutton(label="Include Sideboard", variable=self.include_sideboard)
        self.toggles_menu.add_checkbutton(label="Include Maybeboard", variable=self.include_maybeboard)

        self.toggles_button.config(menu=self.toggles_menu)
        self.toggles_button.pack(side="left", padx=5, pady=2)

        input_frame = tk.Frame(root)
        input_frame.pack(fill='x', padx=5, pady=5)

        missing_frame = tk.Frame(input_frame)
        missing_frame.pack(side='left', padx=5, pady=5, anchor='n')
        tk.Label(missing_frame, text="Missing Cards").pack(anchor='nw')
        self.missing_listbox = tk.Listbox(missing_frame, height=12, width=25)
        self.missing_listbox.pack(side='left', fill='y')
        missing_scroll = ttk.Scrollbar(missing_frame, orient='vertical', command=self.missing_listbox.yview)
        missing_scroll.pack(side='right', fill='y')
        self.missing_listbox.config(yscrollcommand=missing_scroll.set)

        deck_frame = tk.LabelFrame(input_frame, text="Deck Input")
        deck_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        url_frame = tk.Frame(deck_frame)
        url_frame.pack(fill='x', padx=2, pady=2)

        self.url_entry = tk.Entry(url_frame, fg="grey", width=35)
        self.url_entry.insert(0, "Paste a deck link")
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.url_entry.bind("<FocusIn>", self.clear_placeholder)
        self.url_entry.bind("<FocusOut>", self.add_placeholder)

        self.fetch_button = tk.Button(url_frame, text="Fetch", command=self.fetch_deck_from_url)
        self.fetch_button.pack(side="left", padx=2)

        self.last_selected_row = None

        self.save_deck_button = tk.Button(url_frame, text="Save", command=self.save_deck)
        self.save_deck_button.pack(side="left", padx=2)

        self.deck_var = tk.StringVar()
        self.deck_dropdown = ttk.Combobox(url_frame, textvariable=self.deck_var, state="readonly", width=25)
        self.deck_dropdown.pack(side="left", padx=2)
        self.deck_dropdown.bind("<<ComboboxSelected>>", self.load_saved_deck)

        self.deck_dropdown.set("Select saved deck")
        self.deck_dropdown.configure(foreground="grey")
        self.deck_dropdown.bind("<FocusIn>", self.clear_dropdown_placeholder)
        self.deck_dropdown.bind("<FocusOut>", self.add_dropdown_placeholder)

        self.delete_deck_button = tk.Button(url_frame, text="Delete", command=self.delete_deck)
        self.delete_deck_button.pack(side="left", padx=2)

        self.deck_cache = load_deck_cache()
        self.refresh_deck_dropdown()

        self.text_input = tk.Text(deck_frame, height=15, width=60, wrap='word', relief="sunken", borderwidth=2)
        self.text_input.pack(pady=2, padx=2, fill='both', expand=True)

        self.text_input.drop_target_register(DND_FILES)
        self.text_input.dnd_bind("<<Drop>>", self.handle_file_drop)

        control_frame = tk.Frame(root)
        control_frame.pack(fill="x", padx=5, pady=5)

        self.button = tk.Button(control_frame, text='Search Prices', command=self.toggle_search)
        self.button.pack(side="left", padx=5, pady=2)

        self.load_button = tk.Button(control_frame, text='Load File', command=self.load_file)
        self.load_button.pack(side="left", padx=5, pady=2)

        self.save_button = tk.Button(control_frame, text='Save to Excel', command=self.save_to_excel)
        self.save_button.pack(side="left", padx=5, pady=2)

        frame = tk.Frame(root)
        frame.pack(padx=5, pady=5, fill='both', expand=True)

        self.tree = ttk.Treeview(frame, columns=('Card',) + tuple(SCRAPER_CONFIG.keys()) + ('Cheapest',), show='headings')
        for col in self.tree['columns']:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c, False))
            self.tree.column(col, width=100)

        self.context_menu = tk.Menu(self.tree, tearoff=0)

        for source in SCRAPER_CONFIG.keys():
            self.context_menu.add_command(
                label=f"Open from {source}",
                command=lambda s=source: self.open_from_source(s)
            )

        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<ButtonRelease-1>', self.on_click)

        self.progress_label = tk.Label(root, text='')
        self.progress_label.pack(pady=2)
        self.total_label = tk.Label(root, text='Total: AU $0.00', font=('Helvetica', 12, 'bold'))
        self.total_label.pack(pady=5)

        self.tree.bind("<Button-3>", self.show_context_menu) 

        bottom_frame = tk.Frame(root)
        bottom_frame.pack(side="bottom", fill="x", padx=5, pady=5)
        self.open_all_button = tk.Button(bottom_frame, text='Open All Cheapest', command=self.open_all_cheapest)
        self.open_all_button.pack(side="right", padx=5, pady=2)

    def show_context_menu(self, event):
        selected = self.tree.identify_row(event.y)
        if selected:
            self.tree.selection_set(selected) 
            self.context_menu.tk_popup(event.x_root, event.y_root)


    def sort_treeview(self, col, reverse=False):
        rows = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
    
        def try_float(val):
            try:
                return float(val)
            except:
                return val.lower() if isinstance(val, str) else val

        rows.sort(key=lambda t: try_float(t[0]), reverse=reverse)

        for index, (val, k) in enumerate(rows):
            self.tree.move(k, '', index)

        self.tree.heading(col, command=lambda: self.sort_treeview(col, not reverse))

    def clear_placeholder(self, event=None):
        if self.url_entry.get() == "Paste a deck link":
            self.url_entry.delete(0, tk.END)
            self.url_entry.config(fg="black")

    def add_placeholder(self, event=None):
        if not self.url_entry.get().strip():
            self.url_entry.delete(0, tk.END) 
            self.url_entry.insert(0, "Paste a deck link")
            self.url_entry.config(fg="grey")

    def clear_dropdown_placeholder(self, event=None):
        if self.deck_dropdown.get() == "Select saved deck":
            self.deck_dropdown.set("")
            self.deck_dropdown.configure(foreground="black")

    def add_dropdown_placeholder(self, event=None):
        if not self.deck_dropdown.get().strip():
            self.deck_dropdown.set("Select saved deck")
            self.deck_dropdown.configure(foreground="grey")

    def fetch_moxfield_deck(url: str):
        match = re.search(r"/decks/([a-zA-Z0-9\-_]+)", url)
        if not match:
            raise ValueError("Invalid Moxfield URL")
        deck_id = match.group(1)

        api_url = f"https://api.moxfield.com/v2/decks/all/{deck_id}"
        print(f"[DEBUG] Fetching from: {api_url}")

        headers = {
            "User-Agent": "my-mtg-scraper/1.0 (contact: kadenschaedel@gmail.com)"
        }
        r = requests.get(api_url, headers=headers, timeout=15)
        r.raise_for_status()
        data = r.json()

        cards = []
        for section in ("mainboard", "sideboard", "maybeboard"):
            if section in data:
                for card in data[section].values():
                    qty = card["quantity"]
                    name = card["card"]["name"]
                    cards.append((qty, name))
                    print(f"[DEBUG] Parsed {qty}x {name}")
        return cards

    def fetch_deck_from_url(self):
        url = self.url_entry.get().strip()
        if not url or url == "Paste deck link":
            messagebox.showwarning("Input Error", "Please paste a deck link.")
            return

        try:
            import httpx

            if "moxfield.com" in url:
                messagebox.showwarning("Error", "Moxfield not currently supported.")
                return
            else:
                cards = list(mtg_parser.parse_deck(url))

            if not cards:
                messagebox.showerror("Error", "Could not parse decklist from the provided URL.")
                return

            filtered = []
            for c in cards:
                if "sideboard" in c.tags and not self.include_sideboard.get():
                    continue
                if "maybeboard" in c.tags and not self.include_maybeboard.get():
                    continue
                filtered.append(c)

            deck_text = "\n".join([f"{c.quantity} {c.name}" for c in filtered])

            self.text_input.delete('1.0', tk.END)
            self.text_input.insert(tk.END, deck_text)

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to fetch deck:\n{e}")



    def handle_file_drop(self, event):
        path = event.data.strip("{}") 
        if os.path.isfile(path):
            with open(path, 'r', encoding='utf-8') as f:
                self.text_input.delete('1.0', tk.END)
                self.text_input.insert(tk.END, f.read())

    def recalculate_cheapest_prices(self):
        total = 0.0
        missing_cards = []
        for row_id in self.tree.get_children():
            values = self.tree.item(row_id)['values']
            card_name = values[0]
            new_row = [card_name]
            cheapest_price = float('inf')
            cheapest_url = ""
            for source in SCRAPER_CONFIG:
                price_str = self.card_urls.get(card_name, {}).get('Prices', {}).get(source, "--")
                try:
                    price = float(price_str)
                except:
                    price = 0.0
                if self.source_vars[source].get():
                    new_row.append(f"{price:.2f}")
                    if 0 < price < cheapest_price:
                        cheapest_price = price
                        cheapest_url = self.card_urls.get(card_name, {}).get('URLs', {}).get(source, "")
                else:
                    new_row.append("--")
            if cheapest_price == float('inf'):
                cheapest_price = 0.0
            new_row.append(f"{cheapest_price:.2f}")
            self.tree.item(row_id, values=tuple(new_row))
            self.card_urls[card_name]['Cheapest'] = cheapest_url
            total += cheapest_price
            if cheapest_price == 0.0:
                missing_cards.append(card_name)
        self.total_label.config(text=f"Total: AU ${total:.2f}")
        self.missing_listbox.delete(0, tk.END)
        for card in sorted(set(missing_cards)):
            self.missing_listbox.insert(tk.END, card)

    def toggle_search(self):
        if self.button['text'] == 'Search Prices':
            self.button.config(text='Stop')
            self.stop_flag = False
            threading.Thread(target=self.check_prices, daemon=True).start()
        else:
            self.stop_flag = True
            self.button.config(state='disabled')

    def load_file(self):
        filepath = filedialog.askopenfilename(filetypes=[('Text files', '*.txt')])
        if filepath:
            with open(filepath, 'r', encoding='utf-8') as f:
                self.text_input.delete('1.0', tk.END)
                self.text_input.insert(tk.END, f.read())

    def fetch_card_prices_parallel(self, card):
        enabled_sources = {name: cfg['func'] for name, cfg in SCRAPER_CONFIG.items() if self.source_vars[name].get()}
        with ThreadPoolExecutor(max_workers=len(enabled_sources)) as executor:
            futures = {name: executor.submit(func, card) for name, func in enabled_sources.items() if name != "eBay"}
        results = {}
        for name, future in futures.items():
            try:
                result = future.result()
                if isinstance(result, tuple) and len(result) == 3:
                    results[name] = result
                else:
                    results[name] = (0.0, "Invalid result", "")
            except Exception as e:
                print(f"[{name} scrape error]: {e}")
                results[name] = (0.0, "Error", "")
        if "eBay" in enabled_sources:
            time.sleep(0.75)
            results["eBay"] = SCRAPER_CONFIG["eBay"]["func"](card)
        all_sources = SCRAPER_CONFIG.keys()
        prices = []
        urls = []
        display_data = {}
        for name in all_sources:
            if name in results:
                result = results[name]
                price, _, url = result
                prices.append((name, price))
                urls.append((name, url))
                display_data[name] = f"{price:.2f}"
            else:
                display_data[name] = "--"
        cheapest_price = min((p for _, p in prices if p > 0), default=0.0)
        cheapest_url = next((u for n, u in urls if n in results and results[n][0] == cheapest_price), '')
        return (card, display_data, cheapest_price, cheapest_url, results)

    def open_cheapest_from_source(self, source):
        if not self.card_urls:
            messagebox.showinfo("No Results", "Please run a search first.")
            return
        opened = 0
        for card, data in self.card_urls.items():
            cheapest_url = data.get("Cheapest", "")
            source_url = data.get("URLs", {}).get(source, "")
            if source_url and source_url == cheapest_url:
                webbrowser.open_new_tab(source_url)
                opened += 1
        messagebox.showinfo("Done", f"Opened {opened} cheapest links from {source}.")

    def open_from_source(self, source):
        selected = self.tree.selection()
        if not selected:
            return
        row_id = selected[0]
        card_name = self.tree.item(row_id)['values'][0]
        url = self.card_urls.get(card_name, {}).get("URLs", {}).get(source, "")
        if url:
            webbrowser.open_new_tab(url)
        else:
            messagebox.showinfo("No Link", f"No {source} link available for {card_name}.")


    def open_all_cheapest_by_source(self):
        if not self.card_urls:
            messagebox.showinfo("No Results", "Please run a search first.")
            return
        opened = 0
        for source in SCRAPER_CONFIG:
            for card, data in self.card_urls.items():
                url = data.get("URLs", {}).get(source, "")
                price_str = data.get("Prices", {}).get(source, "--")
                try:
                    price = float(price_str)
                except:
                    price = 0.0
                if url and price > 0:
                    webbrowser.open_new_tab(url)
                    opened += 1
        messagebox.showinfo("Done", f"Opened {opened} total links from all sources.")

    def check_prices(self):
        self.tree.delete(*self.tree.get_children())
        self.card_urls.clear()
        self.total_label.config(text='Total: AU $0.00')
        input_text = self.text_input.get('1.0', tk.END)
        cards = parse_decklist_from_input(input_text)
        total = 0.0
        for i, card in enumerate(cards, start=1):
            if self.stop_flag:
                self.progress_label.config(text='Stopped.')
                break
            card, display_data, cheapest, url, results = self.fetch_card_prices_parallel(card)
            row = [card]
            for source in SCRAPER_CONFIG:
                row.append(display_data.get(source, "--"))
            row.append(f"{cheapest:.2f}")
            self.tree.insert('', 'end', values=tuple(row))
            self.card_urls[card] = {
                'Cheapest': url,
                'Prices': {source: display_data.get(source, "--") for source in SCRAPER_CONFIG},
                'URLs': {source: results.get(source, (0.0, "", ""))[2] if results.get(source) else "" for source in SCRAPER_CONFIG}
            }
            total += cheapest
            self.total_label.config(text=f'Total: AU ${total:.2f}')
            self.progress_label.config(text=f'Processing: {i}/{len(cards)}')
            self.root.update_idletasks()
        self.progress_label.config(text='Done' if not self.stop_flag else 'Stopped.')
        self.button.config(text='Search Prices', state='normal')
        self.recalculate_cheapest_prices()

    def open_all_cheapest(self):
        if not self.card_urls:
            messagebox.showinfo("No Results", "Please run a search first.")
            return
        opened = 0
        for card, sources in self.card_urls.items():
            url = sources.get("Cheapest")
            if url:
                webbrowser.open_new_tab(url)
                opened += 1
        messagebox.showinfo("Done", f"Opened {opened} links in your browser.")

    def save_to_excel(self):
        if not self.card_urls:
            messagebox.showinfo("No Data", "You must search prices before saving.")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "MTG Card Prices"
        ws.append(["Card", "Price (AU$)", "Source", "URL"])
        for row_id in self.tree.get_children():
            row = self.tree.item(row_id)['values']
            card = row[0]
            cheapest_price = row[len(SCRAPER_CONFIG) + 1]
            urls = self.card_urls.get(card, {})
            url = urls.get("Cheapest", "")
            source = ""
            if url:
                for name in SCRAPER_CONFIG:
                    if any(substring in url for substring in [
                        "ebay", "moonmtg", "mtgmate", "cardhub", "ggadelaide", "ggmodbury", "goodgames.com.au"
                    ]):
                        if name.lower().replace(" ", "") in url.replace("www.", "").lower():
                            source = name
                            break
            ws.append([card, cheapest_price, source, url])
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"MTG_Price_Report_{timestamp}.xlsx"
        filepath = os.path.join(os.path.expanduser("~/Downloads"), filename)
        try:
            wb.save(filepath)
            messagebox.showinfo("Success", f"Excel file saved to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel file:\n{e}")

    def on_click(self, event):
        selected = self.tree.selection()
        if not selected:
            return

        row_id = selected[0]

        if self.last_selected_row == row_id:
            item = self.tree.item(row_id)
            card_name = item['values'][0]
            urls = self.card_urls.get(card_name, {})
            url = urls.get("Cheapest")
            if url:
                webbrowser.open_new_tab(url)
        else:
            self.last_selected_row = row_id


    def refresh_deck_dropdown(self):
        names = list(self.deck_cache.keys())
        self.deck_dropdown["values"] = names

        if not names:
            self.deck_dropdown.set("Select saved deck")
            self.deck_dropdown.configure(foreground="grey")
        else:
            current = self.deck_var.get()
            if current not in names:
                self.deck_dropdown.set("Select saved deck")
                self.deck_dropdown.configure(foreground="grey")



    def save_deck(self):
        url = self.url_entry.get().strip()
        deck_text = self.text_input.get("1.0", tk.END).strip()
        if not url or not deck_text:
            messagebox.showwarning("Error", "Need both a deck URL and decklist to save.")
            return

        deck_name = None

        if "archidekt.com/decks/" in url:
            try:
                slug = url.rstrip("/").split("/")[-1]
                deck_name = slug.replace("_", " ").title()
            except Exception as e:
                print(f"[Archidekt name parse error] {e}")

        if not deck_name and hasattr(self, "parsed_deck_name") and self.parsed_deck_name:
            deck_name = self.parsed_deck_name

        if not deck_name and "moxfield.com" in url:
            try:
                resp = requests.get(url, timeout=15)
                soup = BeautifulSoup(resp.text, "html.parser")
                tag = soup.select_one("span.deckHeader_deckName__OlKwW")
                if tag:
                    deck_name = tag.get_text(strip=True)
            except Exception as e:
                print(f"[Deck name fetch error] {e}")

        if not deck_name:
            deck_name = f"Deck {len(self.deck_cache)+1}"

        self.deck_cache[deck_name] = {
            "url": url,
            "decklist": deck_text
        }
        save_deck_cache(self.deck_cache)
        self.refresh_deck_dropdown()
        messagebox.showinfo("Saved", f"Deck saved as '{deck_name}'")

    def delete_deck(self):
        name = self.deck_var.get()
        if name in self.deck_cache:
            del self.deck_cache[name]
            save_deck_cache(self.deck_cache)
            self.refresh_deck_dropdown()

            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, "Paste a deck link")
            self.url_entry.config(fg="grey")

            self.text_input.delete("1.0", tk.END)

            self.deck_dropdown.set("Select saved deck")
            self.deck_dropdown.configure(foreground="grey")

            messagebox.showinfo("Deleted", f"Removed deck '{name}'")
        else:
            messagebox.showwarning("Error", "No saved deck selected to delete.")


    def load_saved_deck(self, event=None):
        name = self.deck_var.get()
        if name and name in self.deck_cache:
            data = self.deck_cache[name]

            self.url_entry.delete(0, tk.END)
            self.url_entry.insert(0, data.get("url", ""))
            self.url_entry.config(fg="black")

            self.text_input.delete("1.0", tk.END)
            self.text_input.insert(tk.END, data.get("decklist", ""))

            self.deck_dropdown.configure(foreground="black")


if __name__ == '__main__':
    root = TkinterDnD.Tk()  
    app = MTGScraperGUI(root)
    root.mainloop()


