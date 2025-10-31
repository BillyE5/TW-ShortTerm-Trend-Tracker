import os
import json
import pandas as pd
import pandas_ta as ta
import sys
from datetime import datetime, timedelta, date
from fubon_neo.sdk import FubonSDK
import schedule
import time
import pandas_market_calendars as mcal
import pygame
import msvcrt

# --- 參數設定 ---
DATA_DIRECTORY = r'E:\軟體區\免安裝\MitakeGU\USER\OUT'
BASE_CSV_FILES = [
    # "周轉排行.csv",
    "大單匯集.csv",
    # "大單流入.csv"
]
DATA_ROOT_DIR = r'out_json'

def get_base_watchlist(data_path, base_files, prev_str_yyyymmdd):
    """
    【模組 A】: 結合「手動功課名單」與「大單匯集 CSV」，建立基礎觀察池
    """
    print("--- 名單1:基礎名單 ---")
    
    date_strings = [prev_str_yyyymmdd]
    today_str_yyyymmdd = datetime.now().strftime("%Y%m%d")
    date_strings.append(today_str_yyyymmdd)
    
    # 處理「大單匯集」CSV
    for date_str in date_strings:
        for base_name in base_files:
            file_name = f"{date_str}_{base_name}"
            full_path = os.path.join(data_path, file_name)
            try:
                df = pd.read_csv(
                    full_path, 
                    skiprows=3, 
                    usecols=[2, 3, 5], # C, D, G 欄 (索引 2, 3, 5) 代號、產業別、成交量
                    names=['Symbol', 'Industry', 'Volume'], # 給予對應的欄位名稱
                    encoding='utf-8'
                )
                # --- 基礎篩選 ---
                df.dropna(subset=['Symbol', 'Industry', 'Volume'], inplace=True)
                
                # --- 兩個排除條件 ---
                # 條件一：排除 ETF 和存託憑證
                exclusion_list = ['ETF', '存託憑證', '金融保險', '公司債']
                df = df[~df['Industry'].isin(exclusion_list)]
                
                # 條件二：排除成交量小於 2000 張的股票
                volume_threshold = 2000
                df = df[df['Volume'] >= volume_threshold]
                
                if not df.empty:
                    valid_stocks = df['Symbol'].astype(int).astype(str).str.zfill(4).tolist()
                    all_stocks.update(valid_stocks)

                print(f"  > 從 '{file_name}' 讀取 {len(valid_stocks)} 支股票。")
            except FileNotFoundError:
                print(f"  > 警告：找不到檔案 '{file_name}'，已跳過。")
            except Exception as e:
                print(f"  > 處理 '{file_name}' 時發生錯誤: {e}")
            
    watchlist = sorted(list(all_stocks))
    # print(f"基礎觀察名單建立完畢，共 {len(watchlist)} 支股票。")
    return watchlist

def filter_daytrade_stocks(stock_list, market_info_dict):
    # 當沖檢查 + 參考價
    filtered_list = []
    if len(stock_list) > 0:
        # print("\n--- 正在過濾非當沖標的 ---")
        for stock_id in stock_list:
            try:
                # 這裡的 API 呼叫只在名單建立時執行一次
                ticker_res = restStock.intraday.ticker(symbol=stock_id)
                if ticker_res.get('canBuyDayTrade', False) and float(ticker_res.get('previousClose', 0)) < 500:
                    filtered_list.append(stock_id)

                    # 存參考價
                    market_info_dict[stock_id] = ticker_res.get('referencePrice', None)
            except Exception as e:
                print(f" > 過濾 [{stock_id}] 時發生錯誤: {e}")
                
        # print(f"過濾完畢，剩下 {len(filtered_list)} 支股票(可當沖多 且 股價小於500)。")
    return filtered_list

def get_prev_5mK_data(stock_list, prev_trading_day_obj):
    """【模組 C】: 為指定的股票清單，抓取前一交易日的尾盤 K 棒"""
    print(f"\n開始抓取過去20根 的 5分K 棒資料...")
    
    d_5mK_day_data = {}
    
    for stock_id in stock_list:
        try:
            result = restStock.historical.candles(**{"symbol": stock_id, "timeframe":"5"}) 
            # "from": target_day, "to": target_day,  分K中  時間範圍無效 都預設回傳一個月的資料 由新到舊

            # 檢查回傳結果是否有錯誤碼
            if result.get('statusCode') == 429:
                print(f" > 警告：[{stock_id}] 抓取歷史 K 棒資料失敗，原因：{result.get('message', '未知錯誤')}")
                continue # 跳過這支股票，進行下一輪

            kbars_data = result.get('data', [])

            if len(kbars_data) < 20:
                # 判斷是完全沒有資料，還是資料不足
                if not kbars_data:
                    print(f" > 注意：[{stock_id}] 未抓取到任何歷史 K 棒資料。")
                else:
                    print(f" > 注意：[{stock_id}] 歷史 K 棒資料不足20筆，已忽略。")
                continue

            # 將資料載入 Pandas，準備進行本地端篩選
            df = pd.DataFrame(kbars_data)
            df['datetime'] = pd.to_datetime(df['date']).dt.tz_localize(None)
            
            # 找出前一個交易日
            df_previous_day = df[df['datetime'].dt.date == prev_trading_day_obj]
            if df_previous_day.empty:
                continue

            # 排序並取出最後 20 根
            df_previous_day_sorted = df_previous_day.sort_values(by='datetime', ascending=True)

            # 將 DataFrame 轉換回 List of Dictionaries
            # 使用 .to_dict('records') 方法可以將每一行轉換為一個字典，並組成一個列表。
            last_20_kbars_list = df_previous_day_sorted.tail(20).to_dict('records')

            # key 是股票代號，value 是最前面 20 根 K 棒的 dic
            # 注意：這裡的 value 變成了一個 list
            d_5mK_day_data[stock_id] = last_20_kbars_list

        except Exception as e:
            print(f"  > 抓取 [{stock_id}] 昨日資料時發生錯誤: {e}")

        time.sleep(1) 
    
    print(f"資料準備完畢，共取得 {len(d_5mK_day_data)} 支股票的數據。")
    
    # 3. 在函式結束時，回傳包含所有結果的字典
    return d_5mK_day_data

def find_intraday_strong_stocks():
    """【模組 B】: 在盤中掃描即時排行，找出新的人氣股"""
    print(f"\n--- [{datetime.now().strftime('%H:%M:%S')}] 開始掃描強勢股 ---")

    TOP_N = 80
    try:            
        # --- 步驟 1: 使用 actives 抓取「成交額排行」 ---
        # print("  - 正在抓取成交額排行 (actives)...")
        # trade='value' 代表成交額排行
        actives_by_value = restStock.snapshot.actives(market='TSE', trade='value', type='COMMONSTOCK') 
        otc_actives_by_value = restStock.snapshot.actives(market='OTC', trade='value', type='COMMONSTOCK')

        # --- 步驟 2: 使用 movers 抓取「漲幅排行」 ---
        # print("  - 正在抓取漲幅排行 (movers)...")
        # direction='up', change='percent' 代表漲幅排行
        movers_by_amplitude = restStock.snapshot.movers(market='TSE', direction='up', change='percent', type='COMMONSTOCK', gte=1, lte=9)
        otc_movers_by_amplitude = restStock.snapshot.movers(market='OTC', direction='up', change='percent', type='COMMONSTOCK', gte=1, lte=9)

        # --- 步驟 3: 合併兩份名單，建立初步觀察池 ---
        candidate_symbols = set() # 使用 set 自動過濾重複        

        # 取得當前時間
        now = datetime.now()
        # 根據時間計算動態門檻
        dynamic_threshold = get_dynamic_volume_threshold(now, base_volume=500)

        # 處理 actives 結果
        if actives_by_value.get('data'):
            for stock in actives_by_value['data'][:TOP_N]:
                if stock.get('tradeVolume', 0) > dynamic_threshold: # 以免抓到成交金額大 但張數少的
                    candidate_symbols.add(stock['symbol'])
        if otc_actives_by_value.get('data'):
            for stock in otc_actives_by_value['data'][:TOP_N]:
                if stock.get('tradeVolume', 0) > dynamic_threshold: # 以免抓到成交金額大 但張數少的
                    candidate_symbols.add(stock['symbol'])

        # 處理 movers 結果
        if movers_by_amplitude.get('data'):
            for stock in movers_by_amplitude['data'][:TOP_N]:
                if stock.get('tradeVolume', 0) > dynamic_threshold: # 以免抓到漲幅大 但張數少的
                    candidate_symbols.add(stock['symbol'])
        if otc_movers_by_amplitude.get('data'):
            for stock in otc_movers_by_amplitude['data'][:TOP_N]:
                if stock.get('tradeVolume', 0) > dynamic_threshold: # 以免抓到漲幅大 但張數少的
                    candidate_symbols.add(stock['symbol'])
        
        if not candidate_symbols:
            print("actives/movers 未回傳任何股票。")
            return []
                    
        final_watchlist = list(candidate_symbols)

        print(f"\n篩選完畢，找到 {len(final_watchlist)} 支強勢股。")
        return final_watchlist

    except Exception as e:
        print(f"執行掃描時發生錯誤: {e}")
        return []
    
def get_dynamic_volume_threshold(now, base_volume=500):
    """根據時間計算動態成交量門檻"""
    current_minute = (now.hour - 9) * 60 + now.minute
    
    if current_minute <= 30: # 9:00 - 9:30
        return base_volume
    elif current_minute <= 60: # 9:30 - 10:00
        return base_volume + 600
    elif current_minute <= 90: # 10:00 - 10:30
        return base_volume + 1000
    elif current_minute <= 120: # 10:30 - 11:00
        return base_volume + 1500
    else: # 11:00 以後，可以根據你自己的觀察來設定
        return base_volume + 2000

def run_scan_job(watchlist, d_prev_5mK_data, market_info_dict):
    start_time = time.time()
    print(f"--- [{time.strftime('%H:%M:%S')}] ，開始掃描 ---")
    
    for stock_id in watchlist:
        try:
            # 抓取今日即時 K 棒
            result = restStock.intraday.candles(symbol=stock_id, timeframe=5)
            today_kbars = result.get('data', [])
            
            if not today_kbars:
                continue
                
            # 🔥🔥🔥 拼接昨日與今日的 K 棒資料 🔥🔥🔥
            combined_kbars = d_prev_5mK_data.get(stock_id, []) + today_kbars
            
            df = pd.DataFrame(combined_kbars)
            
            # 將欄位名稱改為 pandas-ta 習慣的格式 (首字母大寫)
            df.rename(columns={'date': 'Date', 'open': 'Open', 'high': 'High', 'low': 'Low', 'close': 'Close', 'volume': 'Volume', 'average': 'Average'}, inplace=True)
            
            if len(df) < 21: # 資料不足以計算 20MA，跳過
                print(f"[{stock_id}] 資料長度不足 ({len(df)} 筆)，跳過分析。")
                continue
            
            # 將 Date 字串轉換為可比較的 datetime 物件
            df['datetime'] = pd.to_datetime(df['Date']).dt.tz_localize(None)

            # 新增 'To' 欄位，對應三竹的時間
            df['To'] = df['datetime'] + pd.Timedelta(minutes=5)

            # --- 計算指標 ---
            
            # 1. 計算 KD(9,3,3)
            # 參數預設就是 (9, 3, 3) (length, smooth_k, smooth_d)
            df.ta.stoch(k=9, d=3, smooth_k=3, append=True) # STOCHk_9_3_3, STOCHd_9_3_3
            
            # 2. 計算 5MA, 10MA, 20MA (用於均線三排)
            df.ta.sma(length=5, append=True)
            df.ta.sma(length=10, append=True)
            df.ta.sma(length=20, append=True)

            # --- 開始判斷訊號條件 ---
            latest = df.iloc[-1]
            previous = df.iloc[-2]

            # 條件 1: KD9 黃金交叉
            is_kd9_golden_cross = False
            
            # 確保指標有計算出來
            k_line = 'STOCHk_9_3_3'
            d_line = 'STOCHd_9_3_3'
            
            if k_line in df.columns and d_line in df.columns:
                # KD黃金交叉判斷邏輯：
                # (前一根 K < 前一根 D) 且 (最新 K > 最新 D)
                # 且最新一根 K < 80 (避免高檔鈍化)
                if (df[k_line].iloc[-2] < df[d_line].iloc[-2]) and \
                   (df[k_line].iloc[-1] > df[d_line].iloc[-1]) and \
                   (df[k_line].iloc[-1] < 80):
                    is_kd9_golden_cross = True
            
            # 條件 2: 均線三排 (多頭排列)
            is_price_above_mas = False
            is_price_above_mas = (latest['Close'] > latest['SMA_5'] and \
                                  latest['SMA_5'] > latest['SMA_10'] and \
                                  latest['SMA_10'] > latest['SMA_20'])

            # 條件 3: 爆量 (最新一根成交量 > 前一根成交量 * 1.2)
            has_attack_volume = False
            if latest['Volume'] > previous['Volume'] * 1.2:
                has_attack_volume = True
            
            # 條件 4: 紅K
            is_red_k = latest['Close'] > latest['Open'] 

            # --- ✅ 最終判斷並發出訊號 ---
            final_signal = is_kd9_golden_cross and is_price_above_mas and has_attack_volume and is_red_k
            
            # 參考價和漲幅判斷邏輯
            reference_price = market_info_dict.get(stock_id)

            if reference_price is not None and final_signal:
                current_price = latest['Close']
                price_change_percent = ((current_price - reference_price) / reference_price) * 100

                signal_text = "KD9黃金交叉 & 均線三排 & 爆量1.2倍 & 紅K棒"
                MAX_PROFITABLE_THRESHOLD = 8.5
                
                if price_change_percent > MAX_PROFITABLE_THRESHOLD:
                    print(f"[❌]【 {stock_id} 】技術面符合 {signal_text}，但漲幅過高 {price_change_percent:.2f}%，已跳過。")
                    continue
                else:
                    if not pygame.mixer.music.get_busy(): # 檢查是否正在播放
                        pygame.mixer.music.play()
                    
                    print(f"🔥🔥【 {stock_id} 】{signal_text} 🔥🔥")
                    print(f"  三竹對應時間: {latest['To']}")
                    print(f"  價格: {latest['Close']}")
                    print("-" * 40)

        except Exception as e:
            print(f"處理 [{stock_id}] 時發生錯誤: {e}")
            print("程式即將終止。")
            sys.exit() # 👈 中斷程式

    end_time = time.time()
    duration = end_time - start_time
    print(f"--- [{time.strftime('%H:%M:%S')}] 本輪掃描結束  總耗時: {duration:.2f} 秒 ---\n")
    
# ==============================================================================
#  主流程
# ==============================================================================
if __name__ == "__main__":

    # --- 登入 API ---
    sdk = None
    # 取得目前這支 Python 腳本所在的資料夾絕對路徑
    # 例如：/你的專案/dayTrade
    script_dir = os.getcwd()
    # script_dir = os.path.dirname(os.path.abspath(__file__))

    # 從腳本路徑再往上一層，找到整個專案的根目錄
    # 例如：/你的專案
    project_root = os.path.dirname(script_dir)

    # 組合出設定檔的完整、絕對路徑
    # 例如：/你的專案/config/config.json
    config_filepath = os.path.join(project_root, 'config', 'config.json')

    # 定義你的提示音檔案路徑
    ALERT_SOUND_FILE = os.path.join(script_dir, 'alert.mp3')
    try:
        pygame.mixer.init()
        pygame.mixer.music.set_volume(0.3) 
        print("Pygame 音訊系統初始化成功！")
        
        pygame.mixer.music.load(ALERT_SOUND_FILE) 
    except Exception as e:
        print(f"警告：音訊系統初始化失敗: {e}")

    try:
        with open(config_filepath, 'r', encoding='utf-8') as f:
            config = json.load(f)
        # print(f"成功從 {config_filepath} 載入設定檔。")
        
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

    # 連結 API Server
    sdk = FubonSDK()
    # 2. 在此處填入您的登入資訊 (請參考官方文件或您的 .pfx 憑證設定)
    accounts = sdk.login(user_id, user_password, cert_path, cert_pass)

    sdk.init_realtime() # 建立行情元件連線
    print("行情元件初始化成功！")
    
    restStock = sdk.marketdata.rest_client.stock
    
    now = datetime.now()
    time_0900 = now.replace(hour=9, minute=00, second=0, microsecond=0)
    time_1325 = now.replace(hour=13, minute=25, second=0, microsecond=0)
    if now < time_0900 or now >= time_1325:
        print("收盤時間，程式將關閉。")
    else:
        # 建立台灣證券交易所 (XTAI) 的日曆
        xtai_calendar = mcal.get_calendar('XTAI')

        # 取得今天的日期 (不含時間)
        today_dt = pd.to_datetime(datetime.now().date())

        # 使用日曆，找出距離今天最近的「前一個交易日」
        # schedule 函式會回傳一個包含所有交易日的 DataFrame
        # 我們取 end_date 為今天，然後找出倒數第二個交易日，就是前一個交易日
        # (如果今天本身是交易日，那今天的日期會是最後一個)
        previous_trading_day = xtai_calendar.schedule(start_date=today_dt - timedelta(days=14), end_date=today_dt).index[-2]
        
        # 將它格式化成 API 需要的字串 "YYYY-MM-DD"
        previousday_str = previous_trading_day.strftime("%Y-%m-%d")
        prev_str_yyyymmdd = previous_trading_day.strftime("%Y%m%d")
        prev_trading_day_obj = previous_trading_day.date()
        
        # print(f"根據台股行事曆，前一個交易日為: {previousday_str}")
        print(f"將會讀取日期為 {previousday_str} 的 CSV 檔案...")

        # --- (模組 A) 建立base觀察名單 ---
        base_watchlist = get_base_watchlist(DATA_DIRECTORY, BASE_CSV_FILES, prev_str_yyyymmdd)
        
        market_info_dict = {}
        # 先排除 不能當沖多的名單
        base_watchlist = filter_daytrade_stocks(base_watchlist, market_info_dict)
    
        # --- 抓取base歷史 K 棒資料 ---
        d_prev_5mK_data = {} # 用來存放歷史資料
        time_1040 = now.replace(hour=10, minute=40, second=0, microsecond=0)

        if now <= time_1040 and len(base_watchlist) > 0:
            print("\n--- 偵測到盤初時段，開始抓取基礎名單的歷史資料 ---")
            d_prev_5mK_data = get_prev_5mK_data(base_watchlist, prev_trading_day_obj)

        # --- 進入盤中監控迴圈 ---
        final_watchlist = base_watchlist.copy()
        strong_stock_scan_done = False
        last_run_minute = -1 
        print("\n系統啟動，進入盤中監控模式...")

        while True:
            #不能刪掉 計時器要用
            now = datetime.now() 
            if now < time_0900 or now >= time_1325:
                print("收盤時間，程式將關閉。")
                break

            # 在盤中任何時間第一次啟動時，都去掃描一次強勢股
            if now.hour >= 9 and now.hour <= 13 and not strong_stock_scan_done:
                print("名單2:掃描盤中即時強勢股。")
                # 掃描盤中即時強勢股
                new_stocks = find_intraday_strong_stocks()
                
                # 排除 不能當沖多的名單
                new_stocks = filter_daytrade_stocks(new_stocks, market_info_dict)

                # 找出「新加入」的股票
                newly_added = set(new_stocks) - set(final_watchlist)
                
                # print(f"\n--- 名單2:發現 {len(newly_added)} 支新強勢股！ ---")
                if newly_added and now <= time_1040:
                    print(f"\n--- 去找20根！ ---")
                    add_prev_5mK_data = get_prev_5mK_data(list(newly_added), prev_trading_day_obj)
                    d_prev_5mK_data.update(add_prev_5mK_data)

                final_watchlist = sorted(list(set(final_watchlist + new_stocks)))
                
                if len(final_watchlist) == 0:
                    print(f"監控名單為 {len(final_watchlist)} 筆，將結束監控。")
                    break
                
                print(f"監控名單長度: {len(final_watchlist)} 支。\n")

                strong_stock_scan_done = True

            # --- 每5分鐘觸發一次訊號掃描 ---
            current_minute = now.minute
            # 條件：確保同分鐘(分鐘數是 5 跟 0 結尾)內不重複執行
            if current_minute % 5 == 0 and current_minute != last_run_minute:
                # (模組 D) 對「最終觀察名單」進行策略訊號掃描
                run_scan_job(final_watchlist, d_prev_5mK_data, market_info_dict)
                
                last_run_minute = current_minute # 更新執行紀錄

            # 讓程式休息一秒，降低 CPU 使用率
            time.sleep(1)

    print("程式已成功結束。")
    print("\n----------------------------------------------------")
    print("程式執行完畢，請按 [Enter] 或 [Esc] 鍵關閉視窗...")

    while True:
        # 檢查是否有鍵盤輸入
        if msvcrt.kbhit():
            # 捕獲單一按鍵（不需按 Enter 即可觸發）
            key = msvcrt.getch()
            
            # Enter 鍵的 ASCII 碼是 b'\r' (Carriage Return)
            # Esc 鍵的 ASCII 碼是 b'\x1b'
            if key == b'\r' or key == b'\x1b':
                # 由於你在程式中使用了 pygame (圖片 image_308809.png 顯示你有初始化 Pygame)
                # 建議在程式結束前，清理 pygame 模組，避免殘留
                try:
                    import pygame
                    pygame.quit()
                except:
                    pass
                
                # 找到指定的按鍵，安全退出程式
                sys.exit(0)
