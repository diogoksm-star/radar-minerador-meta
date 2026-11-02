# -*- coding: utf-8 -*-
import os
import re
import json
import urllib.parse
from datetime import datetime
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

MONTH_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
    "ene": 1, "feb": 2, "may": 5, "ago": 8, "sep": 9, "oct": 10, "dic": 12
}

def parse_meta_date(date_str: str, reference_date: datetime = None):
    if not reference_date:
        reference_date = datetime.now()
    if not date_str:
        return 15, ""
    cleaned = date_str.lower().strip()
    m = re.search(r"(\d+)\s*(?:de\s*)?([a-z]+)\.?\s*(?:de\s*)?(\d{4})", cleaned)
    if m:
        dia = int(m.group(1))
        mes_key = m.group(2)[:3]
        ano = int(m.group(3))
        mes = MONTH_MAP.get(mes_key, 1)
        try:
            dt = datetime(ano, mes, dia)
            diff = (reference_date - dt).days
            return max(diff, 0), dt.strftime("%d/%m/%Y")
        except:
            pass
    return 15, date_str

def extract_real_lp(href: str):
    if not href:
        return ""
    if "l.facebook.com/l.php" in href:
        try:
            parsed = urllib.parse.urlparse(href)
            params = urllib.parse.parse_qs(parsed.query)
            target = params.get("u", [""])[0]
            if target:
                return urllib.parse.unquote(target)
        except:
            pass
    if ("http" in href or "wa.me" in href) and "facebook.com" not in href and "fb.com" not in href:
        return href
    return ""

class MetaAdScraper:
    def __init__(self, country: str = "BR", headless: bool = True):
        self.country = country
        self.headless = headless
        self.chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    def scrape_query(self, query: str, max_scrolls: int = 4):
        encoded_q = urllib.parse.quote(query)
        url = f"https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country={self.country}&q={encoded_q}&sort_data[direction]=desc&sort_data[mode]=relevancy_monthly_grouped"
        
        results = []
        seen_ids = set()

        with sync_playwright() as p:
            launch_args = {"headless": self.headless}
            if os.path.exists(self.chrome_path):
                launch_args["executable_path"] = self.chrome_path

            browser = p.chromium.launch(**launch_args)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
                locale="pt-BR",
                viewport={"width": 1400, "height": 900}
            )
            page = context.new_page()

            try:
                page.goto(url, wait_until="networkidle", timeout=35000)
                page.wait_for_timeout(2500)

                for _ in range(max_scrolls):
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    page.wait_for_timeout(1800)

                html = page.content()
                soup = BeautifulSoup(html, "html.parser")

                id_blocks = soup.find_all(string=re.compile(r"Identifica\w+o da biblioteca:\s*\d+", re.IGNORECASE))
                
                for ib in id_blocks:
                    card = ib.parent
                    for _ in range(8):
                        if card and card.parent:
                            if len(card.parent.find_all(string=re.compile(r"Identifica\w+o da biblioteca:\s*\d+", re.IGNORECASE))) > 1:
                                break
                            card = card.parent
                        else:
                            break

                    card_text = card.get_text(separator="\n").strip()
                    id_m = re.search(r"Identifica\w+o da biblioteca:\s*(\d+)", card_text, re.IGNORECASE)
                    ad_id = id_m.group(1) if id_m else ""
                    if not ad_id or ad_id in seen_ids:
                        continue
                    seen_ids.add(ad_id)

                    date_m = re.search(r"Veicula\w+o iniciada em\s*([^\n\r]+)", card_text, re.IGNORECASE)
                    raw_date = date_m.group(1).strip() if date_m else ""
                    dias_no_ar, data_formatada = parse_meta_date(raw_date)

                    dup_m = re.search(r"(\d+)\s*an[uú]ncios\s*usam\s*esse\s*criativo|usam\s*este\s*criativo\s*e\s*texto\s*em\s*(\d+)\s*an[uú]ncios", card_text, re.IGNORECASE)
                    dups = int(dup_m.group(1) or dup_m.group(2) or 1) if dup_m else 1

                    page_name = ""
                    dest_lp = ""
                    for a in card.find_all("a", href=True):
                        href = a["href"]
                        txt = a.get_text().strip()
                        if "facebook.com/" in href and "ads/library" not in href and "l.facebook.com" not in href and txt and not page_name:
                            page_name = txt
                        lp = extract_real_lp(href)
                        if lp and not dest_lp:
                            dest_lp = lp

                    if not page_name:
                        lines = [l.strip() for l in card_text.split("\n") if l.strip()]
                        for l in lines:
                            if l.lower() not in ["ativo", "inativo", "patrocinado", "identificação da biblioteca", "plataformas"] and len(l) < 50:
                                page_name = l
                                break

                    lines = [l.strip() for l in card_text.split("\n") if l.strip()]
                    copy_lines = []
                    skip = True
                    for l in lines:
                        if "patrocinado" in l.lower():
                            skip = False
                            continue
                        if not skip:
                            if any(term in l.lower() for term in ["saiba mais", "obter oferta", "comprar agora", "ver detalhes", "identificação", "plataformas"]):
                                break
                            copy_lines.append(l)

                    copy_text = " ".join(copy_lines[:6])
                    if not copy_text:
                        copy_text = card_text[:300].replace("\n", " ")

                    results.append({
                        "ad_id": ad_id,
                        "query": query,
                        "ad_library_link": f"https://www.facebook.com/ads/library/?id={ad_id}",
                        "page_name": page_name or f"Anunciante {ad_id}",
                        "start_date": data_formatada or raw_date,
                        "dias_no_ar": dias_no_ar,
                        "duplications": dups,
                        "destination_lp": dest_lp,
                        "copy": copy_text[:400]
                    })
            except Exception as e:
                print(f"[Scraper Error]: {e}")
            finally:
                browser.close()

        return results
