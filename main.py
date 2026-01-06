import sys
import os
import pandas as pd
import gspread
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import re

# ==========================================================
# 1. 設定・ログイン情報
# ==========================================================
LOGIN_URL = "https://dailycheck.tc-extsys.jp/tcrappsweb/web/login/tawLogin.html"
LIST_URL = "https://dailycheck.tc-extsys.jp/tcrappsweb/web/routineStation.html"

# ユーザー情報
USER_ID_1 = "0030"
USER_ID_2 = "927583"
PASSWORD = "Ccj-222223"

# 出力先スプレッドシート
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ge_99NgSbNKQnrrHDMM1wL9n5J2g2mc-xFeMcILCOzo/edit?gid=0#gid=0"
SHEET_TAB_NAME = "stationID"

# 認証キー
SERVICE_ACCOUNT_KEY_FILE = "service_account.json"

# ==========================================================
# 2. 初期化処理
# ==========================================================
if not os.path.exists(SERVICE_ACCOUNT_KEY_FILE):
    print("!! エラー: 認証キーファイル(service_account.json)が見つかりません。")
    print("Secretsの設定(GCP_SA_KEY)が正しいか確認してください。")
    sys.exit(1)

# Google Sheets接続
try:
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_KEY_FILE)
except Exception as e:
    print(f"!! 認証エラー: {e}")
    sys.exit(1)

# ブラウザ設定
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--window-size=1920,1080')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
collected_stations = []

try:
    # ------------------------------------------------------
    # [I] ログイン処理
    # ------------------------------------------------------
    print("\n[I. ログイン処理開始]")
    driver.get(LOGIN_URL)
    sleep(3)

    if "login" in driver.current_url.lower():
        try:
            driver.find_element(By.ID, "cardNo1").send_keys(USER_ID_1)
            driver.find_element(By.ID, "cardNo2").send_keys(USER_ID_2)
            driver.find_element(By.ID, "password").send_keys(PASSWORD)
            driver.find_element(By.ID, "password").send_keys(Keys.RETURN)
            sleep(5)
        except Exception as e:
            print(f"ログイン試行中にエラー発生: {e}")
    else:
        print("既にログイン済み、またはページが異なります。")

    # ------------------------------------------------------
    # [II] リスト収集処理
    # ------------------------------------------------------
    print("\n[II. リスト収集開始]")
    driver.get(LIST_URL)
    sleep(5)

    page_count = 1
    while True:
        print(f"--- ページ {page_count} 解析中 ---")
        
        soup = BeautifulSoup(driver.page_source, "html.parser")
        links = soup.find_all("a", href=True)
        
        current_page_found = 0
        for link in links:
            href = link['href']
            if "stationCd=" in href:
                station_name = link.get_text(strip=True)
                match = re.search(r'stationCd=([0-9a-zA-Z]+)', href)
                station_cd = match.group(1) if match else "Unknown"
                
                # 重複防止
                if not any(d['stationCd'] == station_cd for d in collected_stations):
                    collected_stations.append({
                        "area": "-",
                        "station_name": station_name,
                        "stationCd": station_cd
                    })
                    current_page_found += 1

        print(f"  -> {current_page_found} 件 取得成功")
        
        if current_page_found == 0:
            print("  !! 注意: このページで1件も取れませんでした。")
            break

        # 次ページへ移動ロジック（ID指定版）
        # HTML解析結果: id="assignNextPageBtn" が「次へ」ボタンの実体
        try:
            # 2種類のIDの可能性を考慮（念のため両方チェック）
            next_btn = None
            try:
                next_btn = driver.find_element(By.ID, "assignNextPageBtn")
            except:
                try:
                    next_btn = driver.find_element(By.ID, "allNextPageBtn")
                except:
                    pass

            if next_btn and next_btn.is_displayed():
                # 親要素が disabled クラスを持っていないか確認（終了判定）
                parent_class = next_btn.find_element(By.XPATH, "./..").get_attribute("class")
                if "disabled" in str(parent_class):
                    print("これ以上ページはありません(Disabled)。収集終了。")
                    break
                
                print(f"  [次へ]ボタン(ID:{next_btn.get_attribute('id')})をクリックします")
                driver.execute_script("arguments[0].click();", next_btn)
                sleep(8) # 読み込み待ち（長めに確保）
                page_count += 1
            else:
                print("これ以上ページが見つかりません（ボタンなし）。収集終了。")
                break
                
        except Exception as e:
            print(f"ページ移動処理でエラー: {e}")
            break

    print(f"\n合計 {len(collected_stations)} 件収集完了")

    # ------------------------------------------------------
    # [III] スプレッドシートへ保存
    # ------------------------------------------------------
    if collected_stations:
        print(f"\n[III. スプレッドシート({SHEET_TAB_NAME})へ書き込み]")
        sh = gc.open_by_url(SHEET_URL)
        
        try:
            ws = sh.worksheet(SHEET_TAB_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_TAB_NAME, rows=len(collected_stations)+10, cols=5)

        df_new = pd.DataFrame(collected_stations)
        df_new = df_new[['area', 'station_name', 'stationCd']]
        
        ws.clear()
        ws.update([df_new.columns.values.tolist()] + df_new.values.tolist(), "A1")
        print("書き込み完了！")
    else:
        print("データが0件のため、書き込みをスキップしました。")

except Exception as e:
    print(f"\n!! システムエラー発生: {e}")

finally:
    driver.quit()
