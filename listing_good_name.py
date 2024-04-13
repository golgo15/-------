import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
# import tkinter as tk
# from tkinter import messagebox
import random
import time
# import ctypes


#------------------------------------------------
# 定数
#------------------------------------------------
DEBUG_ENABLE = 0    #0:無効、1:有効
GLOBAL_KEYS = {"yahoo_id"}

#---------------------------------------------------------
#   関数名:read_config()
#   概要:設定ファイルを読み込んで変数に代入する
#   引数:ファイル名
#   戻り値:読み取った内容(key:value)
#---------------------------------------------------------
def read_config(filename):
    config = {}
    with open(filename, 'r') as file:
        for line in file:
            line = line.strip()  # 改行を除去
            if line.startswith('#') or '=' not in line:
                continue  # コメント行またはイコールが含まれない行はスキップ
            key, value = line.split('=', 1)  # イコールで行を分割
            if key.strip() in GLOBAL_KEYS:
                config[key.strip()] = value.strip()  # キーと値を格納
            else:
                print(f"Invalid key: {key.strip()}. Skipping.")
    return config

#---------------------------------------------------------
#   関数名:update_config()
#   概要:設定ファイルを更新する
#   引数:ファイル名
#   戻り値:読み取った内容(key:value)
#---------------------------------------------------------
def update_config(filename, key, new_value):
    updated_lines = []
    found_key = False  # 新しいキーが見つかったかどうかを示すフラグ
    with open(filename, 'r') as file:
        for line in file:
            if line.strip().startswith(key + "="):
                line = f"{key}={new_value}\n"  # 新しい値で行を更新
                found_key = True
            updated_lines.append(line)

    # 新しいキーが見つからなかった場合、ファイルの末尾に追加
    if not found_key:
        updated_lines.append(f"{key}={new_value}\n")

    with open(filename, 'w') as file:
        file.writelines(updated_lines)

#---------------------------------------------------------
#   関数名:write_config()
#   概要:設定ファイルを更新する
#   引数:ファイル名
#   戻り値:読み取った内容(key:value)
#---------------------------------------------------------
def write_config(filename):
    # globalsにあるキーを取得して設定ファイルを更新する
    for key in setting_data:
        value = setting_data[key]  # キーに対応する値を取得
        update_config(filename, key, str(value))  # 設定ファイルを更新

    print("設定ファイルが更新されました。")

#---------------------------------------------------------
#   関数名:DBG_DupDataSet()
#   概要:デバッグ用に重複データをセットする
#---------------------------------------------------------
def DBG_DupDataSet( items ):
    # @@debug
    auction_id = items[0]['data-auction-id']
    auction_id = "4" + auction_id[1:]    #オークションIDは重複しないので先頭を別の文字に変更しておく
    auction_title = items[0]['data-auction-title']
    all_data.append((auction_id, auction_title))


#---------------------------------------------------------
#   関数名:GetRandomTime()
#   概要:指定範囲内での乱数（小数第１位まで）を取得する
#---------------------------------------------------------
def GetRandomTime( startNo, EndNo ):
    # 現在の時間をシードとして設定
    seed = int(time.time())

    # 1.0から3.0までの乱数を生成
    random.seed(seed)
    random_number = random.uniform(startNo, EndNo)

    # 小数を第1位までに制限
    random_time = round(random_number, 1)
    # print(f"待機時間: {random_time:.2f}秒")

    return random_time

#---------------------------------------------------------
#   関数名:AnalizeResponse()
#   概要:取得したhtmlを分析する
#   引数:なし
#   戻り値:次ページの有無（0=次ページなし、1=次ページあり）
#---------------------------------------------------------
def AnalizeResponse(response, all_data):
    if response.status_code == 200:
        html = response.text
        soup = BeautifulSoup(html, "html.parser")

        # data-auction-idとdata-auction-titleを持つ要素を取得
        items = soup.find_all(lambda tag: tag.has_attr('data-auction-id') and tag.has_attr('data-auction-title'))

        # データをall_dataに追加
        for i in range(0, len(items), 2):  # 2つずつ取得するように変更
            auction_id = items[i]['data-auction-id']
            auction_title = items[i]['data-auction-title']
            all_data.append((auction_id, auction_title))

        if DEBUG_ENABLE == 1:
            DBG_DupDataSet(items)  #debug

        # 指定されたセレクターにマッチする要素を取得し、Pager__link--disableが含まれているかチェック
        next_link = soup.select_one("#allContents > div.gv-l-wrapper.gv-l-wrapper--pc.gv-l-wrapper--liquid > main > div > div.gv-l-contentBody > div.gv-l-main > div.Pager > ul > li.Pager__list.Pager__list--next > span.Pager__link--disable")
        if next_link:
            print("データ取得完了")
            return 0

        # 次のページへ
        return 1
    else:
        print("Failed to retrieve data from Yahoo Auctions")
        return 0

#---------------------------------------------------------
#   関数名:DuplicateCheck()
#   概要:重複チェックを行う
#   引数:[i]all_data        チェック対象データ
#        [o]unique_data     チェック済み（重複無し）データ
#        [o]multiple_data   チェック済み（重複）データ
#   戻り値:なし
#---------------------------------------------------------
def DuplicateCheck( all_data, unique_data, multiple_data):
    # 重複を除去する
    # seen = set()
    seen_titles = set()
    for auction_id, auction_title in all_data:
        if auction_title not in seen_titles:    #重複無し
            seen_titles.add(auction_title)
            unique_data.append((auction_id, auction_title))
        else:   #重複あり
            multiple_data.append((auction_id, auction_title))

#---------------------------------------------------------
#   関数名:ExportExcelSheet()
#   概要:指定のエクセルシートに指定データを書き込む
#   引数:[i]unique_data     チェック済み（重複無し）データ
#       
#   戻り値:なし
#---------------------------------------------------------
def ExportExcelSheet( ws, unique_data ):
    ws.append(["Auction ID", "Auction Title", "ITCode"])
    for auction_id, auction_title in unique_data:
        ws.append([auction_id, auction_title, auction_title[-12:]])

#---------------------------------------------------------
#   関数名:OutputResultDuplicate()
#   概要:重複あり時の結果を出力する
#   引数:[i]all_data        チェック対象データ
#        [i]unique_data     チェック済み（重複無し）データ
#        [i]multiple_data   チェック済み（重複）データ
#   戻り値:なし
#---------------------------------------------------------
def OutputResultDuplicate( all_data, unique_data, multiple_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "出品リスト（ストクリ）"
    ExportExcelSheet( ws, unique_data )

    ws_multi = wb.create_sheet("重複出品データ")
    ws_multi.append(["Auction ID", "Auction Title", "ITCode"])
    for auction_id, auction_title in multiple_data:
        ws_multi.append([auction_id, auction_title, auction_title[-12:]])

    filename = "ヤフオク出品リスト_重複あり.xlsx"
    wb.save(filename)
    
    # tkを使うとファイルが大きくなるのでプロンプトで表示に変更
    print("重複出品がありましたのでファイルに保存しました。\n重複出品データシートの商品で、ReCOREと紐づいていない商品をストクリにて削除して下さい。")
    print(f"ストクリの出品リストを実行ファイルの場所に出力します。\nファイル名 => {filename}")
    input("終了します。何かキーを押して下さい")

#---------------------------------------------------------
#   関数名:OutputResultNoDuplicate()
#   概要:重複なし時の結果を出力する
#   引数:[i]all_data        チェック対象データ
#        [i]unique_data     チェック済み（重複無し）データ
#        [i]multiple_data   チェック済み（重複）データ
#   戻り値:なし
#---------------------------------------------------------
def OutputResultNoDuplicate( all_data, unique_data, multiple_data):
    # Excelファイルを作成し、データを書き込む
    wb = Workbook()
    ws = wb.active
    ws.title = "出品リスト（ストクリ）"
    ws.append(["Auction ID", "Auction Title", "ITCode"])
    for auction_id, auction_title in all_data:
        ws.append([auction_id, auction_title, auction_title[-12:]])
    filename = "ヤフオク出品リスト_重複なし.xlsx"
    wb.save(filename)

    # tkを使うとファイルが大きくなるのでプロンプトで表示に変更
    print("重複出品はありませんでした。")
    print(f"ストクリの出品リストを実行ファイルの場所に出力します。\nファイル名 => {filename}")
    input("終了します。何かキーを押して下さい")

#---------------------------------------------------------
#   関数名:OutputResult()
#   概要:結果を出力する
#   引数:[i]all_data        チェック対象データ
#        [i]unique_data     チェック済み（重複無し）データ
#        [i]multiple_data   チェック済み（重複）データ
#   戻り値:なし
#---------------------------------------------------------
def OutputResult( all_data, unique_data, multiple_data):
    # データが重複している場合は別のエクセルファイルに書き出す
    if len(all_data) != len(unique_data):
        OutputResultDuplicate( all_data, unique_data, multiple_data)
    else:
        OutputResultNoDuplicate( all_data, unique_data, multiple_data)


#---------------------------------------------------------
#   関数名:main()
#   概要:メイン処理
#   引数:なし
#   戻り値:なし
#---------------------------------------------------------
# # 白子URL
base_url = "https://auctions.yahoo.co.jp/seller/sngbe73253?sid=sngbe73253&is_postage_mode=1&dest_pref_code=13&b={}&n=100&mode=2"
# 豊川URL
# base_url = "https://auctions.yahoo.co.jp/seller/eyemc48423?sid=sngbe73253&is_postage_mode=1&dest_pref_code=13&b={}&n=100&mode=2"
# ctypes.windll.shcore.SetProcessDpiAwareness(True)

# データを保持するリスト
all_data = []
setting_data = {}
config_file = "setting.ini"     # 設定ファイルのパス

# 設定ファイルを読み込み、設定を取得
config = read_config(config_file)

# 設定情報を変数に代入
for key, value in config.items():
    setting_data[key] = value

# 設定ファイル更新
write_config(config_file)

print('工具買取王国　鈴鹿白子23号店　ヤフオク重複出品チェック開始')
# print('工具買取王国　豊川店　ヤフオク重複出品チェック開始')
start_page = 1
while True:
    time.sleep(GetRandomTime(1.0, 3.0))

    url = base_url.format(start_page)
    response = requests.get(url)
    start_time = time.time()
 
    if 1 == AnalizeResponse(response, all_data):
        start_page += 100
    else:
        break

    end_time = time.time()
    # 処理時間を計算
    elapsed_time = end_time - start_time
    # print(f"処理時間: {elapsed_time:.4f}秒")

# 重複チェックを行う
unique_data = []    # 重複しないデータを格納するリスト
multiple_data = []  # 重複しているデータを格納するリスト
DuplicateCheck( all_data, unique_data, multiple_data)
OutputResult( all_data, unique_data, multiple_data)
