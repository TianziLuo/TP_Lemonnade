from playwright.sync_api import sync_playwright
import time

def teapplix_upload(username, email, password, csv_path):
    chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

    with sync_playwright() as p:
        browser = p.chromium.launch(executable_path=chrome_path, headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # 登录
        page.goto("https://www.teapplix.com/auth/")
        page.fill('input[placeholder="账户名"]', username)
        page.fill('input[placeholder="登录电子邮件"]', email)
        page.fill('input[placeholder="密码"]', password)
        page.click('button.ant-btn-primary')
        page.wait_for_load_state("networkidle")
        print(f"✅ 登录成功：{username}")

        # Inventory
        page.get_by_text("Inventory", exact=True).click()
        page.wait_for_selector('text=Quantity', timeout=10000)
        page.get_by_text("Quantity", exact=True).nth(0).click()
        print("✅ 成功点击 Quantity")

        # Import/Export
        page.get_by_text("Import/Export", exact=True).click()
        page.wait_for_selector("text=Create product automatically", timeout=20000)
        page.get_by_text("Create product automatically", exact=True).click()

        # 上传文件
        page.set_input_files('input[type="file"]', csv_path)
        time.sleep(1)

        # 导入文件
        page.get_by_text("Import CSV", exact=True).click()
        page.wait_for_selector("text=Import CSV", timeout=10000)
        input("🟢 文件上传完成，按 Enter 关闭浏览器...")
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