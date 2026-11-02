# -*- coding: utf-8 -*-
import os
import re
from playwright.sync_api import sync_playwright

class LPDissector:
    def __init__(self, screenshots_dir: str = "output/screenshots"):
        self.screenshots_dir = screenshots_dir
        os.makedirs(self.screenshots_dir, exist_ok=True)
        self.chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    def dissect(self, url: str, offer_name: str = "oferta", timeout_ms: int = 25000):
        safe_name = re.sub(r'[^a-zA-Z0-9]', '_', offer_name)[:25]
        screen_name = f"{safe_name}.png"
        screen_path = os.path.join(self.screenshots_dir, screen_name)

        headline = ""
        checkout = "Não identificado"
        preco = "R$ 27 - R$ 47 (est.)"

        with sync_playwright() as p:
            launch_args = {"headless": True}
            if os.path.exists(self.chrome_path):
                launch_args["executable_path"] = self.chrome_path

            browser = p.chromium.launch(**launch_args)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
                viewport={"width": 1366, "height": 850}
            )
            page = context.new_page()

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)
                page.wait_for_timeout(2500)

                # Headline
                h1 = page.query_selector("h1")
                if h1:
                    headline = h1.inner_text().strip().replace("\n", " ")
                if not headline:
                    h2 = page.query_selector("h2")
                    if h2:
                        headline = h2.inner_text().strip().replace("\n", " ")

                html = page.content()

                # Checkout detection
                if "kiwify.com.br" in html or "pay.kiwify.com.br" in html:
                    checkout = "Kiwify"
                elif "cakto.com.br" in html or "pay.cakto.com.br" in html:
                    checkout = "Cakto"
                elif "hotmart.com" in html or "pay.hotmart.com" in html:
                    checkout = "Hotmart"
                elif "eduzz.com" in html or "sun.eduzz.com" in html:
                    checkout = "Eduzz"
                elif "cartpanda.com" in html:
                    checkout = "Cartpanda"
                elif "kirvano.com" in html:
                    checkout = "Kirvano"
                elif "monetizze.com.br" in html:
                    checkout = "Monetizze"

                # Price detection
                prices = re.findall(r"R\$\s*(\d+[\.,]\d{2})", html)
                if prices:
                    valid_prices = []
                    for pr in prices:
                        val = float(pr.replace(".", "").replace(",", "."))
                        if 9.0 <= val <= 197.0:
                            valid_prices.append(val)
                    if valid_prices:
                        preco_min = min(valid_prices)
                        preco = f"R$ {preco_min:.2f}".replace(".", ",")

                page.screenshot(path=screen_path, full_page=False)

            except Exception as e:
                headline = "Guia digital com acesso imediato após a compra"
            finally:
                browser.close()

        return {
            "url": url,
            "headline": headline[:130] or "Guia digital com entrega imediata",
            "checkout": checkout,
            "preco": preco,
            "screenshot_path": screen_path
        }
