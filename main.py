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
import urllib.parse

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
    # [II] リスト収集処理（まずはURLを集める）
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
                
                # 詳細ページへの絶対URLを作成
                detail_url = urllib.parse.urljoin(driver.current_url, href)

                # 重複防止
                if not any(d['stationCd'] == station_cd for d in collected_stations):
                    collected_stations.append({
                        "area": "-",
                        "station_name": station_name,
                        "stationCd": station_cd,
                        "detail_url": detail_url
                    })
                    current_page_found += 1

        print(f"  -> {current_page_found} 件 取得成功")
        
        if current_page_found == 0:
            print("  !! 注意: このページで1件も取れませんでした。")
            break

        # 次ページへ移動ロジック
        try:
            next_btn = None
            try:
                next_btn = driver.find_element(By.ID, "assignNextPageBtn")
            except:
                try:
                    next_btn = driver.find_element(By.ID, "allNextPageBtn")
                except:
                    pass

            if next_btn and next_btn.is_displayed():
                parent_class = next_btn.find_element(By.XPATH, "./..").get_attribute("class")
                if "disabled" in str(parent_class):
                    print("これ以上ページはありません(Disabled)。")
                    break
                
                print(f"  [次へ]ボタンをクリックします")
                driver.execute_script("arguments[0].click();", next_btn)
                sleep(5)
                page_count += 1
            else:
                print("これ以上ページが見つかりません。")
                break
                
        except Exception as e:
            print(f"ページ移動処理でエラー: {e}")
            break

    print(f"\n合計 {len(collected_stations)} 件の基本情報を取得完了。")

    # ------------------------------------------------------
    # [II-2] 詳細情報の補完
    # ------------------------------------------------------
    print(f"\n[II-2. 詳細ページからエリア情報を取得中...]")
    print("※1件ずつアクセスするため、完了まで数分かかります。")
    
    for i, station in enumerate(collected_stations):
        try:
            target_url = station['detail_url']
            driver.get(target_url)
            sleep(1) # サーバー負荷軽減と読み込み待ち
            
            soup = BeautifulSoup(driver.page_source, "html.parser")
            
            full_address = "-"
            
            # 住所等の項目を探す
            target_th = soup.find(lambda tag: tag.name == "th" and re.search(r'(住所|所在地|エリア|設置場所)', tag.get_text()))
            
            if target_th:
                target_td = target_th.find_next_sibling("td")
                if target_td:
                    full_address = target_td.get_text(strip=True)
            
            # バックアップ：都道府県名を探す
            if full_address == "-":
                 match = soup.find(string=re.compile(r'(都|道|府|県)'))
                 if match and len(match.strip()) < 50:
                     full_address = match.strip()

            # -------------------------------------------
            # 【ここが修正点】住所の整形処理 (〇〇市まで)
            # -------------------------------------------
            if full_address != "-":
                # 1. 都道府県（東京都、北海道、京都府、大阪府、xx県）を削除
                address_no_pref = re.sub(r'^(東京都|北海道|京都府|大阪府|.{2,3}県)', '', full_address)
                
                # 2. 最初の「市」「区」「町」「村」までを抽出
                # 例: 横浜市西区 -> 横浜市
                # 例: 新宿区西新宿 -> 新宿区
                match_city = re.search(r'^(.+?[市区町村])', address_no_pref)
                
                if match_city:
                    station['area'] = match_city.group(1)
                else:
                    # マッチしない場合（海外や特殊な表記）はそのまま入れる
                    station['area'] = address_no_pref
            else:
                station['area'] = "-"
            # -------------------------------------------

            # 進捗表示
            if (i + 1) % 10 == 0:
                print(f"  ... {i + 1}/{len(collected_stations)} 件完了")
            
        except Exception as e:
            print(f"  [{i+1}] 詳細取得エラー: {e}")

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
