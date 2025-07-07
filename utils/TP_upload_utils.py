from playwright.sync_api import sync_playwright
import time

def teapplix_upload(username, email, password, csv_path):
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=chrome_path, headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # Login
        page.goto("https://www.teapplix.com/auth/")
        page.fill('input[placeholder="账户名"]', username)
        page.fill('input[placeholder="登录电子邮件"]', email)
        page.fill('input[placeholder="密码"]', password)
        page.click('button.ant-btn-primary')
        page.wait_for_load_state("networkidle")
        print(f"✅ Logged in: {username}")

        # Inventory
        page.get_by_text("Inventory", exact=True).click()
        page.wait_for_selector('text=Quantity', timeout=10000)
        page.get_by_text("Quantity", exact=True).nth(0).click()
        print("✅ Clicked on 'Quantity'")
     # Import/Export
        page.get_by_text("Import/Export", exact=True).click()
        page.wait_for_selector("text=Create product automatically", timeout=20000)
        page.get_by_text("Create product automatically", exact=True).click()

        # Upload file
        page.set_input_files('input[type="file"]', csv_path)
        time.sleep(1)

        # Import CSV
        page.get_by_text("Import CSV", exact=True).click()
        print("⏳ Please click the 'Close' when finished.")

        # Wait for 'Close' button to disappear after user clicks it
        try:
            page.wait_for_selector('text=Close', state="detached", timeout=120000)
            print("✅ 'Close' button clicked. Closing browser.")
        except:
            print("⚠️ Timeout waiting for 'Close' button to be clicked. Closing browser anyway.")

        browser.close()

r'''
   
if __name__ == "__main__":
    teapplix_upload(
        username="colourtree",
        email="colourtreeusa@gmail.com",
        password="Colourtree168!",
        csv_path=r"C:\Users\monica\Downloads\TP-Upload.csv"
    )
'''
