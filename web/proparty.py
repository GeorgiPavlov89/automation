# proparty.py
from playwright.sync_api import sync_playwright, expect

BASE_URL = "https://portal.registryagency.bg/"
TIMEOUT_MS = 7000  # можеш да го настроиш според нуждите си

def click_when_visible(locator, timeout=TIMEOUT_MS):
    # Изрично изчакване елементът да стане видим, после клик
    expect(locator).to_be_visible(timeout=timeout)
    locator.click()

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # смени на True при нужда
        context = browser.new_context()
        page = context.new_page()

        page.goto(BASE_URL)
        page.wait_for_load_state("domcontentloaded")

        # 1) "Потребител" (бутон)
        click_when_visible(page.get_by_role("button", name="Потребител"))

        # 2) "Вход" (линк)
        click_when_visible(page.get_by_role("link", name="Вход"))

        # 3) "Вход със сертификат" (линк)
        click_when_visible(page.get_by_role("link", name="Вход със сертификат"))

        # 4) Първият линк в банера
        banner_first_link = page.get_by_role("banner").get_by_role("link").first
        click_when_visible(banner_first_link)

        # по желание: изчакай да се приключат заявките след навигация
        page.wait_for_load_state("networkidle")

        # ... тук добави следващи стъпки

        context.close()
        browser.close()

if __name__ == "__main__":
    run()
