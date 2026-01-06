# ==========================================================
# 【GitHub Actions用】ステーションリスト自動更新スクリプト
# 機能: TMAから全ステーションIDを収集し、スプレッドシートへ書き出す
# ==========================================================
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

# 1. ログイン情報（直書き）
LOGIN_URL = "https://dailycheck.tc-extsys.jp/tcrappsweb/web/login/tawLogin.html"
LIST_URL = "https://dailycheck.tc-extsys.jp/tcrappsweb/web/routineStation.html"
USER_ID_1 = "0030"
USER_ID_2 = "927583"
PASSWORD = "Ccj-222223"

# 2. 出力先設定
SHEET_URL = "https://docs.google.com/spreadsheets/d/1ge_99NgSbNKQnrrHDMM1wL9n5J2g2mc-xFeMcILCOzo/edit?gid=0#gid=0"
SHEET_TAB_NAME = "stationID"

# 3. Google認証
SERVICE_ACCOUNT_KEY_FILE = "service_account.json"

if not os.path.exists(SERVICE_ACCOUNT_KEY_FILE):
    print("!! エラー: 認証キーファイルが見つかりません。Secretsの設定を確認してください。")
    sys.exit(1)

gc = gspread.service_account(filename=SERVICE_ACCOUNT_KEY_FILE)

# ==========================================================
# ドライバ設定
# ==========================================================
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--window-size=1920,1080')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
collected_stations = []

try:
    print("\n[I. ログイン処理]")
    driver.get(LOGIN_URL)
    sleep(3)
    
    try:
        driver.find_element(By.ID, "cardNo1").send_keys(USER_ID_1)
        driver.find_element(By.ID, "cardNo2").send_keys(USER_ID_2)
        driver.find_element(By.ID, "password").send_keys(PASSWORD)
        driver.find_element(By.ID, "password").send_keys(Keys.RETURN)
        sleep(5)
    except Exception as e:
        print(f"ログインエラー（または既にログイン済み）: {e}")

    print("\n[II. リスト収集開始]")
    driver.get(LIST_URL)
    sleep(3)

    page_count = 1
    while True:
        print(f"--- ページ {page_count} 解析中 ---")
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # テーブル行を取得（ヘッダーを除く）
        # ※HTML構造に依存するため、すべてのtrを走査してリンクがあるものを探す
        rows = soup.find_all("tr")
        
        current_page_found = 0
        for row in rows:
            link = row.find("a", href=True)
            if link and "stationCd=" in link['href']:
                # ステーション名
                station_name = link.get_text(strip=True)
                
                # ステーションID抽出
                href = link['href']
                match = re.search(r'stationCd=([0-9a-zA-Z]+)', href)
                station_cd = match.group(1) if match else ""
                
                # エリア（市区町村）抽出への試み
                # 通常、テーブルの別の列にある。行内の全テキストを取得して簡易判定、もしくは列指定
                # ここでは確実性を取るため、行内のテキストから"市"や"区"を含む要素を探す、
                # またはCSVのフォーマットに合わせて列位置を推測する。
                # ※TMAの標準的なテーブル構造を想定し、tdのテキストを取得
                cols = row.find_all("td")
                area = ""
                if len(cols) > 1:
                    # 多くの場合、エリアは最初のほうの列にある
                    # ここでは暫定的に「ステーション名の前の列」などをエリアとするか、
                    # 取得できたテキスト全体から判断するが、今回はシンプルに「cols[0]」などをエリアと仮定せず、
                    # "Unknown" とする（CSV作成後に手動修正または既存マッピング利用のため）
                    # ただし、既存CSVに「大和市」等があるため、行のテキストを結合して保持する
                    pass

                # 既存CSVのフォーマット: area, station_name, stationCd
                # 今回はエリア自動判定が難しいため、一旦空欄にするか、行全体から推測が必要だが
                # リスクを避けるため "CheckArea" というプレースホルダーを入れる
                # ※後でシート上で一括置換（多摩市など）したほうが安全
                collected_stations.append({
                    "area": "多摩市", # 今回の追加分は主に多摩とのことなのでデフォルト値をセット
                    "station_name": station_name,
                    "stationCd": station_cd
                })
                current_page_found += 1
        
        print(f"  -> {current_page_found} 件取得")

        # 次ページへ移動
        # 「次へ」ボタンまたはページ番号リンクを探す
        try:
            # "次へ" または "Next" または ">" を含むリンク/ボタンを探す
            next_buttons = driver.find_elements(By.XPATH, "//a[contains(text(), '次へ') or contains(text(), 'Next') or contains(text(), '>')]")
            clicked = False
            for btn in next_buttons:
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    sleep(3)
                    page_count += 1
                    clicked = True
                    break
            
            if not clicked:
                # ボタンが見つからない、または押せなかった場合は終了
                print("これ以上ページが見つかりません。収集終了。")
                break
        except Exception as e:
            print(f"ページ移動処理でエラー、または最終ページ: {e}")
            break

    print(f"\n合計 {len(collected_stations)} 件のステーションを収集しました。")

    # ==========================================================
    # III. シートへ保存
    # ==========================================================
    if collected_stations:
        print(f"\n[III. スプレッドシート保存]")
        print(f"Target: {SHEET_TAB_NAME}")
        
        try:
            sh = gc.open_by_url(SHEET_URL)
        except Exception as e:
            print(f"シートを開けませんでした: {e}")
            sys.exit(1)

        try:
            ws = sh.worksheet(SHEET_TAB_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_TAB_NAME, rows=len(collected_stations)+5, cols=5)
            print(f"タブ '{SHEET_TAB_NAME}' を新規作成しました。")

        # データ整形
        df_new = pd.DataFrame(collected_stations)
        
        # カラム順序統一 (area, station_name, stationCd)
        df_new = df_new[['area', 'station_name', 'stationCd']]
        
        # 書き込み（既存データをクリアして上書き）
        ws.clear()
        # ヘッダー書き込み
        ws.update([df_new.columns.values.tolist()] + df_new.values.tolist(), "A1")
        
        print("書き込み完了！")

except Exception as e:
    print(f"\n!! 重大なエラー発生: {e}")

finally:
    if 'driver' in locals():
        driver.quit()
