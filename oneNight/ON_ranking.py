import os
import sys
import json
import pandas as pd
from datetime import datetime, timedelta
from fubon_neo.sdk import FubonSDK
import pandas_market_calendars as mcal

# version yyyymmdd
__version__ = "20250925"

# --- 設定區 ---
# 根據程式運行環境，動態決定設定檔路徑
if getattr(sys, 'frozen', False):
    # --- 情境 1: 部署 / 分享 (最簡潔的方式) ---
    base_path = os.path.dirname(sys.executable)
elif 'ipykernel' in sys.modules:
    # --- 情境 2: ipynb 腳本 (Jupyter 環境) ---
    # 在 Jupyter 中，os.getcwd() 通常是腳本所在的目錄
    base_path = os.getcwd()
else:
    # --- 情境 3: .py 腳本 (一般 Python 環境) ---
    # .py 腳本會使用 __file__ 變數來取得自身路徑
    base_path = os.path.dirname(os.path.abspath(__file__))

# 每日榜單 輸出資料夾路徑
# 假設 out_oneNight 資料夾也在 .py/.ipynb/.exe 所在的目錄
OUTPUT_DIRECTORY = os.path.join(base_path, 'out_oneNight')

# 設定檔路徑
# config.json 在專案根目錄 (oneNight 的上層)
project_root = os.path.dirname(base_path)
config_path = os.path.join(project_root, 'config', 'config.json')

# --- 設定結束 ---

# --- 判斷今天是否為交易日 ---
def is_trading_day():
    """判斷今天是否為台灣股市的交易日。"""
    xtai_calendar = mcal.get_calendar('XTAI')
    today_dt = datetime.now().date()

    # 取得今天是否為交易日的列表
    valid_trading_days = xtai_calendar.valid_days(start_date=today_dt, end_date=today_dt)

    # 將列表中的日期轉換為不含時區的日期 (remove timezone)
    valid_trading_days_naive = valid_trading_days.tz_convert(None)

    # 將今天的日期轉換成 datetime object
    today_datetime = pd.to_datetime(today_dt)

    # 檢查今天是否在台灣股市的交易日列表中
    return today_datetime in valid_trading_days_naive

def fetch_and_select_stocks():
    """從 API 抓取漲幅榜資料，並根據條件篩選股票。"""
    try:
        print("正在抓取上市及上櫃漲幅榜資料...")
        # 取得上市漲幅榜(COMMONSTOCK為一般股票)
        tse_movers = restStock.snapshot.movers(
            market='TSE', direction='up', change='percent', type='COMMONSTOCK', gte=5
        )['data']
        # 取得上櫃漲幅榜(COMMONSTOCK為一般股票)
        otc_movers = restStock.snapshot.movers(
            market='OTC', direction='up', change='percent', type='COMMONSTOCK', gte=5
        )['data']
        
        # 合併上市櫃資料並轉換成 DataFrame
        all_movers = tse_movers + otc_movers
        df = pd.DataFrame(all_movers)
        
        # 2. 以漲幅欄位 (change_percent) 進行降冪排序
        df.sort_values(by='changePercent', ascending=False, inplace=True)
        # inplace=True 會直接修改 DataFrame，而不是回傳一個新的
    except Exception as e:
        print(f"從 API 取得資料時發生錯誤: {e}")
        return pd.DataFrame()

    # API 欄位名稱與你的 Excel 報表不符，需要重新命名
    df.rename(columns={
        'symbol': '股號',
        'name': '名稱',
        'changePercent': '幅度％',
        'closePrice': '成交',
        'tradeVolume': '成交量',
        'tradeValue': '成交值'
    }, inplace=True)

    condition_gain = df['幅度％'] >= 5 and df['幅度％'] <= 10
    condition_volume = df['成交量'] >= 1500
    condition_value = (df['成交量'] * 1500) * df['成交'] >= 90000000
    # '產業別'不用篩選 "ETF|公司債" 了，已經使用 type = COMMONSTOCK 一般股票
    
    selected_df = df[condition_gain & condition_volume & condition_value].copy()
    
    # 記錄收盤價是否高於當日均價
    selected_df['站上均價'] = ''
    for index, row in selected_df.iterrows():
        symbol = row['股號']
        try:
            day_data = restStock.intraday.quote(symbol=symbol)
            
            if not day_data:
                selected_df.loc[index, '站上均價'] = '無資料'
                continue

            latest_close = day_data['closePrice']
            latest_vwap = day_data['avgPrice']
            
            is_above_vwap = latest_close > latest_vwap  # 站上均價
            
            selected_df.loc[index, '站上均價'] = '是' if is_above_vwap else '否'

        except Exception as e:
            print(f" > 警告：判斷 {symbol} 時發生錯誤: {e}")
            selected_df.loc[index, '站上均價'] = '錯誤'    

    return selected_df

def save_daily_selection(selected_stocks, output_dir):
    if selected_stocks.empty:
        print("今日沒有符合條件的股票。")
        return

    if not os.path.exists(output_dir):
        print(f"建立新資料夾: {output_dir}")
        os.makedirs(output_dir)

    now = datetime.now()
    today_str_ymd = now.strftime('%Y-%m-%d')
    output_df = pd.DataFrame({
        '選股日': today_str_ymd,
        '股票代號': selected_stocks['股號'].astype(int),
        '股票名稱': selected_stocks['名稱'],
        '當日收盤價': selected_stocks['成交'],
        '當日漲幅%': selected_stocks['幅度％'],
        '站上均價': selected_stocks['站上均價'],
        '隔日開盤價': '', # 預留欄位
        '隔日最高價': '', # 預留欄位
        '隔日最低價': '', # 預留欄位
        '隔日收盤價': '', # 預留欄位
        '賣出價格': '',
        '停利/停損': '',
        '損益%': ''
    })
    
    # 如果在中午 12:00 到 下午 1:30 之間執行，則加上 "_中午"
    today_str_filename = now.strftime('%Y%m%d')
    if now.hour <= 13 and now.minute <= 30:
        output_filename = f"{today_str_filename}_漲幅_中午.xlsx"
    else:
        output_filename = f"{today_str_filename}_漲幅.xlsx"
        
        # 2. 備份檔案：用於計算 -1.5% 停損的專屬 exe
        output_filename_15 = f"{today_str_filename}_漲幅_15.xlsx"
        output_filepath_15 = os.path.join(output_dir, output_filename_15)
        output_df.to_excel(output_filepath_15, index=False)
    output_filepath = os.path.join(output_dir, output_filename)
    output_df.to_excel(output_filepath, index=False)
    print(f"榜單已儲存至: {output_filepath}")

if __name__ == "__main__":
    print(f"version : {__version__}")
    if not is_trading_day():
        print(f"今天是 {datetime.now().date().strftime('%Y-%m-%d')}，非台股交易日，程式結束。")
    else:
        print(f"今天是 {datetime.now().date().strftime('%Y-%m-%d')}，台股交易日，開始執行程式。")

        # --- 登入 ---
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 從設定檔中取得登入資訊
            fubon_config = config['fubon_api']
            user_id = fubon_config['id']
            user_password = fubon_config['password']
            cert_path = fubon_config['cert_path']
            cert_pass = fubon_config['cert_pass']

        except FileNotFoundError:
            print("錯誤：找不到 config.json 設定檔！")
            exit()
        except KeyError:
            print("錯誤：config.json 檔案中的 key 不正確！")
            exit()

        sdk = None
        # 連結 API Server
        sdk = FubonSDK()

        accounts = sdk.login(user_id, user_password, cert_path, cert_pass)

        sdk.init_realtime() # 建立行情元件連線
        print("行情元件初始化成功！")
        
        restStock = sdk.marketdata.rest_client.stock

        selected_stocks_today = fetch_and_select_stocks()
        
        # 儲存結果
        save_daily_selection(selected_stocks_today, OUTPUT_DIRECTORY)