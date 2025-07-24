# modules/config.py
import os

class Config:
    DEBUG_MODE = True  # ✅ True 表示调试模式（跳过时间判断和重复写入判断）
    USERNAME = "shihaodong"
    PASSWORD = "12138hh."
    CHROME_PATH = r"E:\chromedriver-win64\chromedriver-win64\chromedriver.exe"

    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")

    DATA_FILE = os.path.join(OUTPUT_DIR, "health_table.xlsx")
    SUMMARY_FILE = os.path.join(OUTPUT_DIR, "hourly_summary_text.xlsx")

    BRAND_ALIASES = {
        "ahold": "Ahold",
        "delhaize": "Delhaize",
        "delhaize-prd": "Delhaize-prd",
        "delhaize-acc": "delhaize-acc",
        "aldi-sued": "aldi-sued",
        "aldi-usa": "aldi-usa",
        "aldi-au": "aldi-au",
        "wow": "wow",
        "ahold-ab": "ahold-ab",
        "leroymerlin": "leroymerlin",
        "ab": "AB"
    }
