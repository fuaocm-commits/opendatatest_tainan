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
    level=logging. INFO #設定記錄層級
    fo

)