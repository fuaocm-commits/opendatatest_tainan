# 載入所需套件工具
from sqlalchemy.engine import create
import os # 處理路徑
import logging  # 紀錄log
import requests as rq # 向server提出請求
import pandas as pd # 處理資料
from sqlalchemy import create_engine, text # DB engine
import pymysql # 連線mysql
import openpyxl # excel驅動
import urllib3 # 處理url
from dotenv import load_dotenv # 抓取.env資料

# 1. 載入.env
load_dotenv()

# 2. 設定log 機制(純寫到file & 輸出到終端機上)
logging.basicConfig(
    level=logging.INFO, # 設定記錄層級
    format='%(asctime)s [%(levelname)s] %(filename)s(行:%(lineno)d: %(message)s)',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("scraper.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# 設定變數
DB_HOST = os.environ.get("DB_HOST", "localhost")
DB_USER = os.environ.get("DB_USER", "peter")
DB_PASSWORD = os.environ.get("DB_PASSWORD", "[PASSWORD]")
DB_PORT = int(os.environ.get("DB_PORT", 3306))
DB_NAME = os.environ.get("DB_NAME", "tainan")
DB_CHARSET = os.environ.get("DB_CHARSET", "utf8mb4")

API_URL = os.environ.get("API_URL")
EXCEL_FILENAME = os.environ.get("EXCEL_FILENAME", "tainan_house.xlsx")

# 透過 API 抓取 opendata
def fetch_data(api_url, params=None):
    # 檢查api_url
    if not api_url:
        logging.error("API URL 為空!")
        return None

    try:
        logging.info(f"正在從API get data...")
        # get opendata by API
        res = rq.get(api_url, params=params, verify=False)

        # 檢查回傳狀態
        if res.ok: # 200
            # 轉換資料型態 json -> python
            data = res.json()
            logging.info(f"成功取得資料, 共 {len(data.get('data', []))} 筆")
            return data
        else:
            logging.warning(f"回應資料錯誤, {res.status_code}")
            return None
    except rq.exceptions.SSLError as ssl_err:
        logging.error(f"SSL 認證錯誤:{ssl_err}")
        return None
    except Exception as error:
        logging.error(f"執行擷取open data時, 發生錯誤:{error}")
        return None

    #以上為0903上課內容
def save_to_excel(data,filename):
    try:
        #pandas針對不同維度的資料型態不同 :一維->series 二維:dataframe
        df=pd.DataFrame(data)
        df.to_excel(filename,index=False,engine="openpyxl") 
        #false原因是因為不需要dataframe之下的序號
        #透過openxl這個驅動開啟excel
        logging.info(f"資料寫入成功{filename}")
    except Exception as e:
        #exc_info:把trace印出來
        logging.error(f"儲存excel時發生錯誤:{e}",exc_info=True)#exc_info=True這是參數；官方參數

def save_to_mysql(data): #py檔案要操作DB資料須執行步驟: 1.建立通道 2.通道中用SQL語法溝通
    #建立通道
    conn=None
    #SQL指令物件
    cursor=None

    try:
        #將資料轉換成pandas的資料型態,excel(2D):dataframe
        df = pd.DataFrame(data)

        # 1.建立通道
        logging.info(f"連線至MySQL({DB_HOST}:{DB_PORT})檢查資料庫狀態 ... ")
        conn=pymysql.connect(
            host=DB_HOST,#主機在哪裡
            user=DB_USER,
            password=DB_PASSWORD,
            port=DB_PORT,
            charset=DB_CHARSET #字元集
        )
        cursor=conn.cursor()

        #建立資料庫
        cursor.execute(f"CREATE DATABASE IF NOT EXISTS {DB_NAME} CHARACTER SET {DB_CHARSET} COLLATE utf8mb4_unicode_ci;")
        conn.commit()
        cursor.close()
        conn.close()
        conn=None
        cursor=None

        #使用sqlalchemy進行table建立與資料寫入
        db_uri = f"mysql+pymysq1://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset={DB_CHARSET}"
        engine =create_engine(db_uri)

        # 建立資料表 table
        # 產生通道
        with engine.begin() as sql_conn:
        # 傳遞sq1指令
            sql_conn.execute(text("""
                CREATE TABLE IF NOT EXISTS house (
                Seq BIGINT,
                鄉鎮市區別 TEXT,
                區段數合計 TEXT,
                一般區段數 TEXT,
                繁榮街道路線價區段數 TEXT,
                一般路線價區段數 TEXT,
                一般區段價最高 TEXT,
                一般區段價最低 TEXT,
                最高繁榮街道路線價 TEXT,
                最低繁榮街道路線價 TEXT,
                最高一般路線價 TEXT,
                最低一般路線價 TEXT,
                最高宗地地價 TEXT
                )ENGINE=InnoDB DEFAULT CHARSET=utf8 COLLATE=utf8_unicode_ci;"""))

            #清空table舊資料
            sql_conn.excute(text("TRUNCATE TABLE house;"))
        logging.info("MysQl 資料表已建立並清空舊資料")

        #寫入資料
        logging.info("爭將資料寫入 SQL table...中")
        df.tp_sql(
            name="house",
            con=engine,
            if_exist="append",
            index=False,
            chunksize=1000,
            method="multi"
        )
        logging.info("資料成功寫入mysql!")
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        logging.error(f"MySQL #fFi#li: {e}", exc_info=True)
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def read_from_excel(filename):
    try:
        db_uri = f"mysql+pymysq1://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}?charset={DB_CHARSET}"
        engine=create_engine(db_uri)
        df=pd.read_excel(filename,engine="openpyxl")
        logging.info(f"成功讀取excel,共(len(df))筆資料")
        return df
    except Exception as e :
        logging.error (f"讀取excel失敗!{e}")
        return None

if __name__=="__main__":
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    data = fetch_data(API_URL)

if data and isinstance(data, dict) and 'data' in data:
    records = data['data']

    if records:
        logging.info(f"準備處理{len(records)}筆記錄")

        # 儲存資料
        save_to_excel(records, EXCEL_FILENAME)
        save_to_mysql(records)

