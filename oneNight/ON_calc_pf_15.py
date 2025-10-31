import pandas as pd
import os
import json
from datetime import datetime, timedelta
from openpyxl import load_workbook
import pandas_market_calendars as mcal
import sys

# 匯入 SDK Library
from fubon_neo.sdk import FubonSDK, Order
from fubon_neo.constant import TimeInForce, OrderType, PriceType, MarketType, BSAction

# version yyyymmdd
__version__ = "20250925"

STOP_PROFIT_PERCENT = 1.5 # 設定停利點為 1.5%
STOP_LOSS_PERCENT = -1.5  # 設定停損點為 -1.5%
TRANSACTION_COST_PERCENT = 0.6 # 交易成本

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

# 台灣證券交易所 (XTAI) 的日曆
xtai_calendar = mcal.get_calendar('XTAI')
# --- 設定結束 ---

# --- 判斷今天是否為交易日 ---
def is_trading_day():
    """判斷今天是否為台灣股市的交易日。"""
    today_dt = datetime.now().date()

    # 取得今天是否為交易日的列表
    valid_trading_days = xtai_calendar.valid_days(start_date=today_dt, end_date=today_dt)

    # 將列表中的日期轉換為不含時區的日期 (remove timezone)
    valid_trading_days_naive = valid_trading_days.tz_convert(None)

    # 將今天的日期轉換成 datetime object
    today_datetime = pd.to_datetime(today_dt)

    # 檢查今天是否在台灣股市的交易日列表中
    return today_datetime in valid_trading_days_naive

def get_intraday_minute_kbars(stock_id):
    """
    接收一個股票代號和日期，去 API 請求當天所有的 5分K 資料。

    Args:
        stock_id (str): 股票代號，例如 '2330'。

    Returns:
        list[dict]: 包含當天所有分鐘 K 棒的列表，
                    如果失敗則返回 None。
    """
    try:        
        reststock = sdk.marketdata.rest_client.stock 
        candles = reststock.intraday.candles(symbol=stock_id, timeframe=5)
        today_kbars = candles.get('data', [])

        return today_kbars
    except Exception as e:
        print(f"錯誤：無法取得 {stock_id} 的分鐘K棒資料. 原因: {e}")
        return None

def run_high_precision_backtest(purchase_price, stock_id):
    """
    取得分鐘K棒，模擬盤中走勢，找出最先觸發的事件。
    """

    take_profit_price = purchase_price * (1 + (STOP_PROFIT_PERCENT / 100))
    stop_loss_price = purchase_price * (1 + (STOP_LOSS_PERCENT / 100))

    # 取得當天所有的分鐘K棒
    minute_kbars = get_intraday_minute_kbars(stock_id)
    if not minute_kbars: return None
    
    # --- 從分鐘K棒中，預先計算出當日的完整 OHLC ---
    day2_open = minute_kbars[0].get('open', purchase_price)
    day2_high = max(k.get('high', 0) for k in minute_kbars)
    day2_low = min(k.get('low', float('inf')) for k in minute_kbars)
    day2_close = minute_kbars[-1].get('close', purchase_price)
    
    daily_ohlc_data = {
        'open': day2_open,
        'high': day2_high,
        'low': day2_low,
        'close': day2_close
    }

    # --- 判斷出場邏輯 ---
    result = ''
    pnl_percent = 0
    sell_price = 0

    # 檢查開盤價是否向上跳空缺口 直接超過停利點，以開盤價賣出
    if day2_open >= take_profit_price:
        result = '停利'
        sell_price = day2_open # 賣出價為開盤價
        pnl_percent = (((sell_price - purchase_price) / purchase_price) * 100) - TRANSACTION_COST_PERCENT
    else:
        # 如果沒有開盤跳空，才進入盤中模擬
        for kbar in minute_kbars:
            # 檢查這根K棒的最低價，是否"先"觸及停損點 (風險優先 當作先觸發停損)
            if kbar.get('low', float('inf')) <= stop_loss_price:
                result = '停損'
                sell_price = stop_loss_price # 賣出價為停損價
                pnl_percent = STOP_LOSS_PERCENT - TRANSACTION_COST_PERCENT
                break
            
            # 再檢查這根K棒的最高價，是否觸及停利點
            if kbar.get('high', 0) >= take_profit_price:
                result = '停利'
                sell_price = take_profit_price # 賣出價為停利價
                pnl_percent = STOP_PROFIT_PERCENT - TRANSACTION_COST_PERCENT
                break

    # 如果迴圈跑完，result 依然是空的，代表盤中都沒觸發
    if not result:
        sell_price = day2_close # 賣出價為收盤價
        
        pnl_percent = (((sell_price - purchase_price) / purchase_price) * 100) - TRANSACTION_COST_PERCENT
        if pnl_percent > 0:
            result = '停利'
        else:
            result = '停損'

    return {
        'result': result,
        'pnl_percent': pnl_percent, 
        'sell_price': sell_price,
        'daily_ohlc': daily_ohlc_data
    }

def generate_summary_and_save(df, filepath):
    """
    直接使用已算好的「損益%」欄位來計算總金額，確保邏輯一致。
    """
    # --- 設定區 ---
    SHARES_PER_TRADE = 1000
    COST_FACTOR = 1.006  # 將所有買賣成本統一估算為 0.6%，並計入買入成本

    # --- 1. 篩選出已完成的交易 ---
    completed_trades = df[df['賣出價格'].notna()].copy()

    if completed_trades.empty:
        print("沒有已完成的交易，無法產生績效報告。")
        df.to_excel(filepath, index=False)
        return

    # --- 2. 用價格進行加總計算 ---
    
    # 計算買入股票的總價值 (未計費用)
    total_purchase_value = completed_trades['當日收盤價'].sum() * SHARES_PER_TRADE
    
    # 計算賣出股票的總價值 (未計費用)
    total_sell_value = completed_trades['賣出價格'].sum() * SHARES_PER_TRADE

    # 總投入成本 (大成本)，*1.006
    total_investment_cost = total_purchase_value * COST_FACTOR
    
    # 總盈虧 (元)：總賣出收入 - 總投入成本
    total_net_profit = total_sell_value - total_investment_cost
    
    # 總績效 ROI (%)
    total_roi = (total_net_profit / total_investment_cost) * 100 if total_investment_cost > 0 else 0

    # --- 3. 計算停利/停損次數與勝率 ---
    total_trades = len(completed_trades)
    take_profits = completed_trades[completed_trades['停利/停損'] == '停利'].shape[0]
    stop_losses = total_trades - take_profits
    win_rate = (take_profits / total_trades) * 100 if total_trades > 0 else 0

    # --- 4. 建立績效報告 ---
    summary_data = {
        '項目': ['總交易張數', '停利次數', '停損次數', '勝率 (%)',
                 '總投入成本(含費用)', '總盈虧 (元)', '總績效 (ROI %)'],
        '數值': [total_trades, take_profits, stop_losses, f"{win_rate:.2f}",
                 f"{total_investment_cost:,.0f} 元", f"{total_net_profit:,.0f} 元", f"{total_roi:.2f}"]
    }
    summary_df = pd.DataFrame(summary_data)
    
    # 將結果寫入 Excel
    try:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='TradeLog')
            df.update(completed_trades)
            summary_df.to_excel(writer, sheet_name='TradeLog', startrow=0, startcol=14, index=False)
        print("\n績效計算完成！")
    except Exception as e:
        print(f"錯誤：寫入 Excel 檔案時發生錯誤: {e}")

def calculate_performance(data_dir):
    """
    績效計算
    """

    # 取得今天的日期 (不含時間)
    today_dt00 = pd.to_datetime(datetime.now().date())

    # 使用日曆，找出距離今天最近的「前一個交易日」
    #    schedule 函式會回傳一個包含所有交易日的 DataFrame
    #    我們取 end_date 為今天，然後找出倒數第二個交易日，就是前一個交易日
    #    (如果今天本身是交易日，那今天的日期會是最後一個)
    previous_trading_day = xtai_calendar.schedule(start_date=today_dt00 - timedelta(days=14), end_date=today_dt00).index[-2]
    
    # 組合出前一個交易日的檔案路徑
    previous_day_filename = f"{previous_trading_day.strftime('%Y%m%d')}_漲幅.xlsx"
    
    # 取得日期字串
    day_str = previous_trading_day.strftime('%Y%m%d')

    # 判斷要讀取哪個檔案
    # 💡 注意：這裡假設你的 SL15 備份檔是為了測試 -1.5% 的績效
    if STOP_LOSS_PERCENT == -1.5:  
        # -1.5%，讀 _15 檔案
        previous_day_filename = f"{day_str}_漲幅_15.xlsx"
    else:
        # -2.0%
        previous_day_filename = f"{day_str}_漲幅.xlsx"

    log_filepath = os.path.join(data_dir, previous_day_filename)

    try:
        df = pd.read_excel(log_filepath)
        if any(col.startswith('Unnamed') for col in df.columns):
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        if '損益%' not in df.columns: df['損益%'] = pd.NA
    except FileNotFoundError:
        print(f"錯誤：找不到昨天的選股檔案 '{log_filepath}'。")
        return
    to_calculate_mask = df['損益%'].isna()
    if not to_calculate_mask.any():
        print("檔案中所有股票皆已計算過績效。")
        generate_summary_and_save(df, log_filepath)
        return
        
    print(f"找到 {to_calculate_mask.sum()} 筆股票需要計算績效...")

    for index, row in df[to_calculate_mask].iterrows():
        pnl_percent = 0
        result = ''

        stock_id = str(row['股票代號'])
        purchase_price = row['當日收盤價']
        
        # yesterday_date_str = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        backtest_result = run_high_precision_backtest(purchase_price, stock_id)
        
        # 如果回測成功，才寫入結果
        if result is not None and pnl_percent is not None:
            df.loc[index, '停利/停損'] = result
            df.loc[index, '損益%'] = round(pnl_percent, 2)

        if backtest_result:
            ohlc = backtest_result['daily_ohlc']
                        
            df.loc[index, '隔日開盤價'] = ohlc['open']
            df.loc[index, '隔日最高價'] = ohlc['high']
            df.loc[index, '隔日最低價'] = ohlc['low']
            df.loc[index, '隔日收盤價'] = ohlc['close']
            df.loc[index, '賣出價格'] = round(backtest_result['sell_price'], 2)
            df.loc[index, '停利/停損'] = backtest_result['result']
            df.loc[index, '損益%'] = round(backtest_result['pnl_percent'], 2)

    generate_summary_and_save(df, log_filepath)

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
        # 2. 在此處填入您的登入資訊 (請參考官方文件或您的 .pfx 憑證設定)
        accounts = sdk.login(user_id, user_password, cert_path, cert_pass)
        print("Fubon SDK 初始化完畢！")

        sdk.init_realtime() # 建立行情元件連線

        # --- 資料查詢區 ---
        # 建立行情查詢 WebAPI 連線 Object Instance
        restStock = sdk.marketdata.rest_client.stock

        calculate_performance(OUTPUT_DIRECTORY)
