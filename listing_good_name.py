import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
# import tkinter as tk
# from tkinter import messagebox
import random
import time
import ctypes

# # 白子URL
base_url = "https://auctions.yahoo.co.jp/seller/sngbe73253?sid=sngbe73253&is_postage_mode=1&dest_pref_code=13&b={}&n=100&mode=2"
# 豊川URL
# base_url = "https://auctions.yahoo.co.jp/seller/eyemc48423?sid=sngbe73253&is_postage_mode=1&dest_pref_code=13&b={}&n=100&mode=2"
ctypes.windll.shcore.SetProcessDpiAwareness(True)

# データを保持するリスト
all_data = []

print('工具買取王国　鈴鹿白子23号店　ヤフオク重複出品チェック開始')
# print('工具買取王国　豊川店　ヤフオク重複出品チェック開始')
start_page = 1
while True:
    # 現在の時間をシードとして設定
    seed = int(time.time())

    # 1.0から3.0までの乱数を生成
    random.seed(seed)
    random_number = random.uniform(1.0, 3.0)

    # 小数を第1位までに制限
    wait_time = round(random_number, 1)
    # print(f"待機時間: {wait_time:.2f}秒")
    time.sleep(wait_time)

    url = base_url.format(start_page)
    response = requests.get(url)
    start_time = time.time()
 
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
        
        # @@debug
        # auction_id = items[0]['data-auction-id']
        # auction_title = items[0]['data-auction-title']
        # all_data.append((auction_id, auction_title))

        # 指定されたセレクターにマッチする要素を取得し、Pager__link--disableが含まれているかチェック
        next_link = soup.select_one("#allContents > div.gv-l-wrapper.gv-l-wrapper--pc.gv-l-wrapper--liquid > main > div > div.gv-l-contentBody > div.gv-l-main > div.Pager > ul > li.Pager__list.Pager__list--next > span.Pager__link--disable")
        if next_link:
            print("データ取得完了")
            break

        # 次のページへ
        start_page += 100
    else:
        print("Failed to retrieve data from Yahoo Auctions")
        break
    end_time = time.time()
    # 処理時間を計算
    elapsed_time = end_time - start_time
    # print(f"処理時間: {elapsed_time:.4f}秒")

# 重複をチェックし、重複しないデータを格納するリスト
unique_data = []
# 重複しているデータを格納するリスト
multiple_data = []

# 重複を除去する
# seen = set()
seen_titles = set()
for auction_id, auction_title in all_data:
    # if (auction_id, auction_title) not in seen:
    #     seen.add((auction_id, auction_title))
    #     unique_data.append((auction_id, auction_title))
    # else:
    #     multiple_data.append((auction_id,auction_title))

    if auction_title not in seen_titles:
        seen_titles.add(auction_title)
        unique_data.append((auction_id, auction_title))
    else:
        multiple_data.append((auction_id, auction_title))



# データが重複している場合は別のエクセルファイルに書き出す
if len(all_data) != len(unique_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "出品リスト（ストクリ）"
    ws.append(["Auction ID", "Auction Title", "ITCode"])
    for auction_id, auction_title in unique_data:
        ws.append([auction_id, auction_title, auction_title[-12:]])

    ws_multi = wb.create_sheet("重複出品データ")
    ws_multi.append(["Auction ID", "Auction Title", "ITCode"])
    for auction_id, auction_title in multiple_data:
        ws_multi.append([auction_id, auction_title, auction_title[-12:]])

    # @@debug
    # ws_all = wb.create_sheet("全データ")
    # ws_all.append(["Auction ID", "Auction Title", "ITCode"])
    # for auction_id, auction_title in all_data:
    #     ws_all.append([auction_id, auction_title, auction_title[-12:]])

    filename = "ヤフオク出品リスト_重複あり.xlsx"
    wb.save(filename)
    # print(f"Unique data has been exported to '{filename}'")
    
    # tkを使うとファイルが大きくなるのでプロンプトで表示に変更
    print("重複出品がありましたのでファイルに保存しました。\n重複出品データシートの商品で、ReCOREと紐づいていない商品をストクリにて削除して下さい。")
    print(f"ストクリの出品リストを実行ファイルの場所に出力します。\nファイル名 => {filename}")
    input("終了します。何かキーを押して下さい")
    # # ポップアップでメッセージを表示
    # root = tk.Tk()
    # root.withdraw()  # メインウィンドウを非表示にする
    # root.lift()
    # messagebox.showerror("同一IT出品検出！", "重複出品がありましたのでファイルに保存しました。\n重複出品データシートの商品で、ReCOREと紐づいていない商品を\nストクリにて削除して下さい。")
else:
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
    print(f"ストクリの出品リストを実行ファイルの場所に出力します。ファイル名 => {filename}")
    input("終了します。何かキーを押して下さい")
    # # ポップアップでメッセージを表示
    # root = tk.Tk()
    # root.withdraw()  # メインウィンドウを非表示にする
    # root.lift()
    # messagebox.showinfo("重複出品なし", "重複出品はありませんでした。")
