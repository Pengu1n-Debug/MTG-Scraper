import httpx
import threading
import time
import webbrowser
import os
import datetime
import sys
from urllib.parse import urlencode, quote, quote_plus
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

SORTING_MAP = {'best_match': 12}
session = httpx.Client(headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'en-US,en;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br'
}, http2=True, follow_redirects=True, timeout=30.0)

EBAY_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/123.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.ebay.com.au/",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "DNT": "1",
}

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

def fetch_mtgmate_price(card_name):
    from playwright.sync_api import sync_playwright
    from urllib.parse import quote
    import re
    from bs4 import BeautifulSoup
    import random
    import time

    url = f"https://www.mtgmate.com.au/cards/{quote(card_name)}"
    print(f"\n[MTGMate] Fetching: {url}")

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--disable-web-security',
                    '--disable-features=VizDisplayCompositor',
                    '--no-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu',
                    '--disable-extensions',
                    '--no-first-run',
                    '--disable-default-apps',
                    '--disable-background-networking',
                    '--disable-background-timer-throttling',
                    '--disable-renderer-backgrounding'
                ]
            )
            
            context = browser.new_context(
                viewport={"width": random.randint(1200, 1920), "height": random.randint(800, 1080)},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                locale="en-AU",
                timezone_id="Australia/Sydney",
                extra_http_headers={
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                    'Accept-Language': 'en-AU,en;q=0.9,en-US;q=0.8',
                    'Accept-Encoding': 'gzip, deflate, br',
                    'Cache-Control': 'max-age=0',
                    'DNT': '1',
                    'Connection': 'keep-alive',
                    'Upgrade-Insecure-Requests': '1',
                    'Sec-Fetch-Dest': 'document',
                    'Sec-Fetch-Mode': 'navigate',
                    'Sec-Fetch-Site': 'none',
                    'Sec-Fetch-User': '?1',
                    'sec-ch-ua': '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
                    'sec-ch-ua-mobile': '?0',
                    'sec-ch-ua-platform': '"Windows"'
                }
            )
            
            page = context.new_page()
            
            page.add_init_script("""
                // Remove webdriver property
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined,
                });
                
                // Remove chrome automation indicators
                delete window.chrome;
                
                // Mock permissions API
                Object.defineProperty(navigator, 'permissions', {
                    get: () => ({
                        query: () => Promise.resolve({ state: 'granted' })
                    })
                });
                
                // Add realistic plugins
                Object.defineProperty(navigator, 'plugins', {
                    get: () => [1, 2, 3, 4, 5]
                });
                
                // Mock languages
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['en-AU', 'en', 'en-US']
                });
                
                // Override automation detection
                window.navigator.chrome = {
                    runtime: {}
                };
                
                // Remove automation flags
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => false,
                });
            """)
            
            time.sleep(random.uniform(1.5, 3.5))
            
            page.goto(url, timeout=30000, wait_until='domcontentloaded')
            
            page.wait_for_timeout(random.randint(2000, 4000))
            
            page.evaluate("""
                window.scrollTo(0, Math.floor(Math.random() * 200));
                setTimeout(() => window.scrollTo(0, 0), 500);
            """)
            
            page.wait_for_timeout(random.randint(1000, 2000))
            
            html = page.content()
            browser.close()

        soup = BeautifulSoup(html, "html.parser")

        page_text = soup.get_text()

        matches = re.findall(r"\$\d+(?:\.\d{2})?", page_text)
        print(f"[MTGMate] Found price strings: {matches}")

        prices = [float(p.strip('$').replace(',', '')) for p in matches]

        if not prices:
            print("[MTGMate] No price matches found.")
            return (0.0, 'Not found', '')

        cheapest = min(prices)
        print(f"[MTGMate] Cheapest price: ${cheapest:.2f}")
        return (cheapest, f"{card_name} (MTGMate)", url)

    except Exception as e:
        print(f"[MTGMate ERROR] {e}")
        return (0.0, 'Error', '')

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

def scrape_cardhub(card_name):
    def normalize(text):
        text = text.lower()
        text = re.sub(r"[^a-z0-9\s-]", "", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    target = normalize(card_name)
    target_first = target.split()[0] if target else ""

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
            title_first = title_norm.split()[0] if title_norm else ""

            print(f"[CardHub] #{idx} Title: {title} | Price: {price}")
            if title_first != target_first:
                print(f"[CardHub] Skipping: '{title_first}' != '{target_first}'")
                continue

            link_tag = title_div.find_parent("a")
            link = link_tag["href"] if link_tag else ""
            if link and not link.startswith("http"):
                link = "https://thecardhubaustralia.com.au" + link

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

def parse_ebay_search(response):
    previews = []
    
    # Try multiple parsing strategies
    try:
        # Strategy 1: Original Parsel approach
        sel = Selector(response.text)
        best_selling_boxes = sel.xpath(
            '//*[*[h2[contains(text(),"Best selling products")]]]//li[contains(@class, "s-item")]'
        )
        best_selling_html_set = set([b.get() for b in best_selling_boxes])
        
        items_found = sel.css('.srp-results li.s-item')
        print(f"[eBay Parser] Strategy 1: Found {len(items_found)} items with '.srp-results li.s-item'")
        
        for box in items_found:
            if box.get() in best_selling_html_set:
                continue
            css = lambda css: box.css(css).get('') or None
            css_float = (
                lambda css: float(box.css(css).re_first(r'(\d+\.*\d*)', default='0.0'))
                if box.css(css)
                else 0.0
            )
            href = box.css('a::attr(href)').get()
            if not href or not href.startswith("http"):
                continue
            price = css_float('.s-item__price::text')
            if price == 0.0 or price == 20.0:
                continue
            shipping = css_float('.s-item__shipping::text')
            total = price + shipping
            title = css('.s-item__title span::text')
            if not title:
                continue
            title_lower = title.lower()

            if any(bad in title_lower for bad in ['art card', 'art series', 'display commander']):
                continue
            if any(bulk in title_lower for bulk in ['singles', 'choose your card', 'pick your card', 'select card']):
                continue
            if all(good not in title_lower for good in ['mtg', 'magic']):
                continue

            item = {
                'title': title,
                'price': price,
                'shipping': shipping,
                'total': total,
                'url': href
            }
            previews.append(item)
            
        if previews:
            return previews
            
    except Exception as e:
        print(f"[eBay Parser] Strategy 1 failed: {e}")
    
    # Strategy 2: Try alternative selectors
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Try different item selectors
        alternative_selectors = [
            'li.s-item',
            'div[data-testid="item"]',
            '.it-ttl',
            '.lvtitle',
            '.itemContainer'
        ]
        
        for selector in alternative_selectors:
            items = soup.select(selector)
            print(f"[eBay Parser] Strategy 2: Found {len(items)} items with '{selector}'")
            
            if items:
                for item in items[:10]:  # Limit to first 10 items for processing
                    try:
                        # Try to extract title
                        title_elem = item.select_one('h3, .it-ttl, .s-item__title, [data-testid="title"]')
                        if not title_elem:
                            continue
                        title = title_elem.get_text(strip=True)
                        
                        # Try to extract price
                        price_elem = item.select_one('.s-item__price, .notranslate, .u-flL, .fee, .bin')
                        if not price_elem:
                            continue
                        price_text = price_elem.get_text(strip=True)
                        
                        # Extract numeric price
                        import re
                        price_match = re.search(r'AU?\$?([\d,]+\.?\d*)', price_text)
                        if not price_match:
                            continue
                        price = float(price_match.group(1).replace(',', ''))
                        
                        if price == 0.0 or price == 20.0:
                            continue
                        
                        # Try to extract URL
                        link_elem = item.find('a', href=True) or item.select_one('[data-testid="link"]')
                        if not link_elem:
                            continue
                        href = link_elem.get('href', '')
                        
                        if not href.startswith('http'):
                            href = 'https://www.ebay.com.au' + href
                        
                        title_lower = title.lower()
                        
                        # Apply filters
                        if any(bad in title_lower for bad in ['art card', 'art series', 'display commander']):
                            continue
                        if any(bulk in title_lower for bulk in ['singles', 'choose your card', 'pick your card', 'select card']):
                            continue
                        if all(good not in title_lower for good in ['mtg', 'magic']):
                            continue
                        
                        preview_item = {
                            'title': title,
                            'price': price,
                            'shipping': 0.0,  # Default shipping
                            'total': price,
                            'url': href
                        }
                        previews.append(preview_item)
                        
                    except Exception as item_e:
                        print(f"[eBay Parser] Error parsing item: {item_e}")
                        continue
                
                if previews:
                    print(f"[eBay Parser] Strategy 2 successful: Found {len(previews)} valid items")
                    return previews
                    
    except Exception as e:
        print(f"[eBay Parser] Strategy 2 failed: {e}")
    
    print("[eBay Parser] All parsing strategies failed")
    return previews


def scrape_ebay(cardname):
    import random
    import time
    
    url = 'https://www.ebay.com.au/sch/i.html?' + urlencode({
        '_nkw': cardname,
        '_sacat': 0,
        '_ipg': 60,
        '_sop': SORTING_MAP['best_match'],
        '_pgn': 1,
        'LH_BIN': 1,
    })
    
    print(f"\n[eBay] Fetching: {url}")
    
    enhanced_headers = {
        'User-Agent': random.choice([
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15'
        ]),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-AU,en;q=0.9,en-US;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'DNT': '1',
        'Cache-Control': 'max-age=0',
        'Pragma': 'no-cache'
    }
    
    try:
        time.sleep(random.uniform(1, 3))
        
        print("[eBay] Trying enhanced session method...")
        enhanced_session = httpx.Client(
            headers=enhanced_headers,
            http2=True, 
            follow_redirects=True, 
            timeout=30.0
        )
        
        response = enhanced_session.get(url)
        enhanced_session.close()
        
        # Check for challenge page indicators
        challenge_indicators = [
            'checking your browser',
            'pardon our interruption',
            'challenge',
            'captcha',
            'blocked'
        ]
        
        response_lower = response.text.lower()
        is_challenge_page = any(indicator in response_lower for indicator in challenge_indicators)
        
        if is_challenge_page:
            print("[eBay] Challenge page detected, trying alternative approach...")
            # Try to extract the actual destination URL from the challenge page
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Look for the form with the destination URL
            form = soup.find('form', {'id': 'destForm'})
            if form and form.get('action'):
                dest_url = form.get('action')
                print(f"[eBay] Found redirect URL: {dest_url}")
                
                # Wait a bit and try the actual destination
                time.sleep(random.uniform(3, 6))
                
                try:
                    # Create a new session for the destination request
                    dest_session = httpx.Client(
                        headers=enhanced_headers,
                        http2=True, 
                        follow_redirects=True, 
                        timeout=30.0
                    )
                    dest_response = dest_session.get(dest_url)
                    dest_session.close()
                    
                    if not any(indicator in dest_response.text.lower() for indicator in challenge_indicators):
                        print(f"[eBay] Successfully accessed destination page, length: {len(dest_response.text)}")
                        results = parse_ebay_search(dest_response)
                        print(f"[eBay] Parsed {len(results)} results from destination")
                        
                        if results:
                            best = min(results, key=lambda x: x['total'])
                            return (best['total'], best['title'], best['url'])
                except Exception as dest_e:
                    print(f"[eBay] Error fetching destination: {dest_e}")
            
            # If all else fails, try a simplified search approach
            print("[eBay] Trying simplified search...")
            return scrape_ebay_simple(cardname)
        
        print(f"[eBay] Enhanced session success, response length: {len(response.text)}")
        
        results = parse_ebay_search(response)
        print(f"[eBay] Parsed {len(results)} results")
        
        if not results:
            return (0.0, 'Not found', '')
        best = min(results, key=lambda x: x['total'])
        return (best['total'], best['title'], best['url'])
        
    except Exception as e:
        print(f"[eBay ERROR] {e}")
        return (0.0, 'Error', '')


def scrape_ebay_simple(cardname):
    """Simplified eBay scraper as final fallback"""
    print(f"[eBay Simple] eBay has implemented strong bot protection")
    print(f"[eBay Simple] Temporarily returning 'temporarily unavailable' for {cardname}")
    print(f"[eBay Simple] eBay scraping requires more advanced techniques (e.g., residential proxies, CAPTCHA solving)")
    
    # For now, return a special status indicating the service is temporarily unavailable
    # This is more honest than returning "Not found" when we actually can't access the site
    return (0.0, 'Temporarily unavailable (eBay bot protection)', '')


def scrape_ebay_playwright(cardname):
    """Fallback eBay scraper using Playwright for challenge pages"""
    import random
    import time
    from playwright.sync_api import sync_playwright
    
    url = 'https://www.ebay.com.au/sch/i.html?' + urlencode({
        '_nkw': cardname,
        '_sacat': 0,
        '_ipg': 60,
        '_sop': SORTING_MAP['best_match'],
        '_pgn': 1,
        'LH_BIN': 1,
    })
    
    print(f"[eBay Playwright] Fetching: {url}")
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=True,
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--disable-web-security',
                    '--disable-features=VizDisplayCompositor',
                    '--no-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-gpu',
                    '--disable-extensions',
                    '--no-first-run',
                    '--disable-default-apps'
                ]
            )
            
            context = browser.new_context(
                viewport={"width": random.randint(1200, 1920), "height": random.randint(800, 1080)},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                locale="en-AU",
                timezone_id="Australia/Sydney",
                extra_http_headers={
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                    'Accept-Language': 'en-AU,en;q=0.9,en-US;q=0.8',
                    'Accept-Encoding': 'gzip, deflate, br',
                    'Cache-Control': 'max-age=0',
                    'DNT': '1',
                    'Connection': 'keep-alive',
                    'Upgrade-Insecure-Requests': '1'
                }
            )
            
            page = context.new_page()
            
            # Add stealth scripts
            page.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined,
                });
                
                delete window.chrome;
                
                Object.defineProperty(navigator, 'permissions', {
                    get: () => ({
                        query: () => Promise.resolve({ state: 'granted' })
                    })
                });
                
                Object.defineProperty(navigator, 'plugins', {
                    get: () => [1, 2, 3, 4, 5]
                });
                
                Object.defineProperty(navigator, 'languages', {
                    get: () => ['en-AU', 'en', 'en-US']
                });
                
                window.navigator.chrome = {
                    runtime: {}
                };
            """)
            
            time.sleep(random.uniform(1, 3))
            
            # Navigate to the page with more realistic behavior
            print("[eBay Playwright] Navigating to search page...")
            page.goto(url, timeout=45000, wait_until='networkidle')
            
            # Add some realistic mouse movements and scrolling
            page.evaluate("""
                // Simulate some realistic user behavior
                window.scrollTo(0, 100);
                setTimeout(() => window.scrollTo(0, 300), 1000);
                setTimeout(() => window.scrollTo(0, 0), 2000);
            """)
            
            # Wait for potential challenge page to resolve
            initial_wait = random.randint(8000, 12000)
            print(f"[eBay Playwright] Initial wait: {initial_wait}ms")
            page.wait_for_timeout(initial_wait)
            
            # Check if we're still on a challenge page
            page_content = page.content()
            challenge_keywords = ['checking your browser', 'pardon our interruption', 'challenge', 'reference id']
            
            if any(keyword in page_content.lower() for keyword in challenge_keywords):
                print("[eBay Playwright] Challenge page detected, attempting bypass...")
                
                # Try clicking or interacting if there are interactive elements
                try:
                    # Look for any buttons or interactive elements that might help
                    button_selectors = ['button', '[role="button"]', 'input[type="submit"]']
                    for selector in button_selectors:
                        elements = page.query_selector_all(selector)
                        if elements:
                            print(f"[eBay Playwright] Found {len(elements)} {selector} elements")
                            break
                except:
                    pass
                
                # Wait longer for automatic redirection
                extended_wait = random.randint(15000, 25000)
                print(f"[eBay Playwright] Extended wait: {extended_wait}ms")
                page.wait_for_timeout(extended_wait)
                page_content = page.content()
                
                # Final check
                if any(keyword in page_content.lower() for keyword in challenge_keywords):
                    print("[eBay Playwright] Still on challenge page after extended wait")
                else:
                    print("[eBay Playwright] Challenge page resolved!")
            
            browser.close()
            
            # Create a mock response object for parse_ebay_search
            class MockResponse:
                def __init__(self, text):
                    self.text = text
            
            mock_response = MockResponse(page_content)
            results = parse_ebay_search(mock_response)
            print(f"[eBay Playwright] Parsed {len(results)} results")
            
            if not results:
                return (0.0, 'Not found', '')
            
            best = min(results, key=lambda x: x['total'])
            return (best['total'], best['title'], best['url'])
            
    except Exception as e:
        print(f"[eBay Playwright ERROR] {e}")
        return (0.0, 'Error', '')


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

SCRAPER_CONFIG = {
    "eBay": {"enabled": True, "func": scrape_ebay},
    "MoonMTG": {"enabled": True, "func": scrape_moonmtg}, 
    "MTGMate": {"enabled": True, "func": fetch_mtgmate_price},
    "CardHub": {"enabled": True, "func": scrape_cardhub},
    "GGAustralia": {"enabled": True, "func": scrape_ggaustralia},
    "GGModbury": {"enabled": True, "func": scrape_ggmodbury},
    "GGAdelaide": {"enabled": True, "func": scrape_ggadelaide},
}

SOURCE_TO_COLUMN = {
    "eBay": "eBay",
    "MoonMTG": "Moon",
    "MTGMate": "MTGMate",
    "CardHub": "CardHub",
    "GGAustralia": "GGTCG",
    "GGModbury": "GGModbury",
    "GGAdelaide": "GoodGames",
}

class MTGScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title(
            "MTG Price Checker (eBay • MoonMTG • MTGMate • CardHub • GGAustralia • GGAdelaide • GGModbury)"
        )
        self.card_urls = {}
        self.stop_flag = False

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

        text_frame = tk.Frame(input_frame)
        text_frame.pack(side='left', fill='both', expand=True)

        tk.Label(text_frame, text='Enter Deck List (or load from file)').pack(anchor='w')
        self.text_input = tk.Text(text_frame, height=10, width=60, wrap='word')
        self.text_input.pack(pady=2, padx=2, fill='both', expand=True)

        dropdown_frame = tk.LabelFrame(input_frame, text="Open Cheapest Options")
        dropdown_frame.pack(side='right', fill='y', padx=5)

        self.open_all_button = tk.Button(dropdown_frame, text='Open All Cheapest', command=self.open_all_cheapest)
        self.open_all_button.pack(padx=4, pady=2, fill='x')

        for source in SCRAPER_CONFIG:
            btn = tk.Button(dropdown_frame, text=f"From {source}",
                            command=lambda s=source: self.open_cheapest_from_source(s))
            btn.pack(padx=4, pady=1, fill='x')

        self.open_all_sources_button = tk.Button(dropdown_frame, text='All from All Sources',
                                                 command=self.open_all_cheapest_by_source)
        self.open_all_sources_button.pack(padx=4, pady=4, fill='x')


        button_frame = tk.Frame(root)
        button_frame.pack(pady=5)
        self.button = tk.Button(button_frame, text='Search Prices', command=self.toggle_search)
        self.button.pack(side=tk.LEFT, padx=5)
        self.load_button = tk.Button(button_frame, text='Load File', command=self.load_file)
        self.load_button.pack(side=tk.LEFT, padx=5)
        self.save_button = tk.Button(button_frame, text='Save to Excel', command=self.save_to_excel)
        self.save_button.pack(side=tk.LEFT, padx=5)

        frame = tk.Frame(root)
        frame.pack(padx=5, pady=5, fill='both', expand=True)

        self.source_vars = {}
        toggle_frame = tk.Frame(root)
        toggle_frame.pack(pady=5)

        for source in SCRAPER_CONFIG:
            var = tk.BooleanVar(value=SCRAPER_CONFIG[source]['enabled'])
            cb = tk.Checkbutton(toggle_frame, text=source, variable=var, command=self.recalculate_cheapest_prices)
            cb.pack(side=tk.LEFT, padx=3)
            self.source_vars[source] = var

        self.tree = ttk.Treeview(
            frame,
            columns=('Card',) + tuple(SCRAPER_CONFIG.keys()) + ('Cheapest',),
            show='headings'
        )
        for col in self.tree['columns']:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<ButtonRelease-1>', self.on_click)
        self.progress_label = tk.Label(root, text='')
        self.progress_label.pack(pady=2)
        self.total_label = tk.Label(root, text='Total: AU $0.00', font=('Helvetica', 12, 'bold'))
        self.total_label.pack(pady=5)

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
        enabled_sources = {
            name: cfg['func']
            for name, cfg in SCRAPER_CONFIG.items()
            if self.source_vars[name].get()
        }
    
        with ThreadPoolExecutor(max_workers=len(enabled_sources)) as executor:
            futures = {
                name: executor.submit(func, card)
                for name, func in enabled_sources.items()
                if name != "eBay"
            }

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
        if selected:
            item = self.tree.item(selected[0])
            card_name = item['values'][0]
            urls = self.card_urls.get(card_name, {})
            url = urls.get("Cheapest")
            if url:
                webbrowser.open_new_tab(url)


if __name__ == '__main__':
    root = tk.Tk()
    app = MTGScraperGUI(root)
    root.mainloop()


