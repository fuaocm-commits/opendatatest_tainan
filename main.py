#載入所需套件工具
import os #處理路徑
import logging #紀錄log
import requests as rq #向server提出請求
import pandas as pandas #處理資料
from sqlalchemy import create_engine, text #DB的engine
import pymysql #連線mysql
import openpyxl #excel驅動
import urllib3 #處理url
from dotenv import load_dotenv #處理environment 

#1.載入 .env
load_dotenv()

#2. 設定log 機制(純寫到file &輸出到終端機上)
logging.basicConfig(
    level=logging. INFO, # 設定記錄層級
    format='%(asctime)s [%(lenelname)s] %(filename)s(f]:%(lineno)d: %(message)s)',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("scraper.log", encoding="utf-8"),
        logging.StreamHandler()
        ]
)

#設定變數
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_USER = os.environ.get("DB_USER", "peter")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "[PASSWORD]")
DB_PORT = int(os.environ.get("DB_PORT", 3306))
DB_NAME = os.environ.get("DB_NAME", "tainan")
DB_CHARSET = os.environ.get("DB_CHARSET", "utf8mb4")

API_URL = os.environ.get("API_URL")
EXCEL_FILENAME = os.environ.get("EXCEL_FILENAME", "tainan_house.xlsx")
