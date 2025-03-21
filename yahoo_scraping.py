import requests
from bs4 import BeautifulSoup
import xlwings as xw
import os

def scrape_yahoo_news():
    # YahooニュースのURL
    url = "https://news.yahoo.co.jp/"
    
    # Webページを取得
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "html.parser")

    # ニュースの見出しを取得（CSSセレクタを使用）
    headlines = [h.text for h in soup.select("a[href*='/pickup/']")]

    file_name = "スクレイピング結果.xlsx"
    
    # ファイルが存在しない場合は新規作成
    if not os.path.exists(file_name):
        wb = xw.Book()  # 新規ブックを作成
        wb.save(file_name)  # ファイルを保存
        wb.close()

    # Excelに書き込む
    wb = xw.Book(file_name)  # 既存のExcelファイルを開く（なければ新規作成）
    sheet = wb.sheets[0]  # シートを指定

    # ヘッダーを書き込み
    sheet.range("A1").value = ["見出し"]

    # データを書き込み
    sheet.range("A2").value = [[h] for h in headlines]

    # 保存して閉じる
    wb.save()
    wb.close()

    print("スクレイピング完了！Excelに保存しました。")

# 関数を実行
scrape_yahoo_news()
