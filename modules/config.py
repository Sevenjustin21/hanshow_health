# modules/config.py
import os
from dotenv import load_dotenv
from pathlib import Path


load_dotenv(dotenv_path=Path(__file__).resolve().parent.parent / ".env") # 加载根目录的 .env 文件


class Config:
    DEBUG_MODE = False
    USERNAME = os.getenv("MY_APP_USERNAME")
    PASSWORD = os.getenv("MY_APP_PASSWORD")
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
