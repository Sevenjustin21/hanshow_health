# main.py
from modules.config import Config
from modules.login_handler import LoginHandler
from modules.extractor import DataExtractor
from modules.utils import TimeUtils

if __name__ == "__main__":
    config = Config()
    LoginHandler(config).scrape()
    DataExtractor(config, TimeUtils()).extract_and_write()
