# modules/login_handler.py
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from openpyxl import Workbook
import os

class LoginHandler:
    def __init__(self, config):
        self.config = config

    def scrape(self):
        service = Service(self.config.CHROME_PATH)
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        driver = webdriver.Chrome(service=service, options=options)

        try:
            driver.get("https://syswatch.hanshowcloud.com/health/view/global")
            time.sleep(2)

            driver.find_element(By.NAME, "username").send_keys(self.config.USERNAME)
            driver.find_element(By.NAME, "password").send_keys(self.config.PASSWORD)
            driver.find_element(By.XPATH, "//button[@type='submit']").click()
            time.sleep(5)

            driver.get("https://syswatch.hanshowcloud.com/health/view/global")
            time.sleep(5)

            rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            data = [[col.text.strip() for col in row.find_elements(By.TAG_NAME, "td")] for row in rows if row]

            os.makedirs(self.config.OUTPUT_DIR, exist_ok=True)
            wb = Workbook()
            ws = wb.active
            ws.append([
                "客户号", "PS", "EW", "TEMPL", "门店对接", "AP", "当日对接", "当日更新",
                "PS更新数量", "数据错误", "失败标签", "更新时间", "操作"
            ])
            for row in data:
                ws.append(row)
            wb.save(self.config.DATA_FILE)
            print(f"✅ 已抓取数据并保存至 {self.config.DATA_FILE}")
        finally:
            driver.quit()
