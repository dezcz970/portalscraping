import time
from selenium import webdriver
import pandas as pd
from bs4 import BeautifulSoup
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import smtplib, ssl
from email.mime.text import MIMEText


def scraping():
    # 空のデータフレームを定義
    dfs = pd.DataFrame(index=[], columns=['タイトル', '対象', '掲示期間', '発信元', 'サブタイトル', '本文'])

    while True:
        soup_a= BeautifulSoup(browser.page_source, 'html.parser')
        tables_a = soup_a.find_all('table')
        dfs_a = pd.read_html(str(tables_a))[0]
        dfs_a.columns = ['タイトル', '対象', '掲示期間', '発信元']

        # 一行目削除
        dfs_a.drop(0, inplace=True)

        # 中身の情報を取得
        subtitle_list_a = []
        subject_list_a = []
        for i in range(100):
            try:
                #タイトルをクリック
                table = browser.find_element_by_tag_name('table')
                a_tag = table.find_elements_by_tag_name('a')
                a_tag[i].click()
                subtitle_a = browser.find_element_by_xpath('//*[@id="content"]/div/table/tbody/tr[3]/td[2]/span')
                subtitle_list_a.append(subtitle_a.text)
                subject_a = browser.find_element_by_css_selector('table')
                subject_list_a.append(subject_a.text)

                 #一覧へ戻る
                input = browser.find_elements_by_tag_name('input')
                input[-2].click()

            except:
                break
        # listをseriesにする
        subtitle_series_a = pd.Series(subtitle_list_a)
        subject_series_a = pd.Series(subject_list_a)

        # seriesのindexを1からにする
        subtitle_series_a.index = subtitle_series_a.index + 1
        subject_series_a.index = subject_series_a.index + 1

        dfs_a['サブタイトル'] = subtitle_series_a
        dfs_a['本文'] = subject_series_a
        dfs = pd.concat([dfs,dfs_a], ignore_index= True)

        try:
            browser.find_element_by_link_text('next≫').click()

        except:
            break
    return dfs
    time.sleep(2)


def send_line(nothing):
    nothing_number = len(nothing.index)

    for x in range(nothing_number):
        nothing_subtitle = nothing.iloc[x,4]
        nothing_subject = nothing.iloc[x,5]
        nothing_target = nothing.iloc[x,1]
        nothing_source = nothing.iloc[x,3]
        nothing_term = nothing.iloc[x,2]

        # headers
        access_token = '1369X7sb6ZYQEhXC7m3By7xNSw7KBrFst5CmPUqEGEL' # 発行されたトークンへ置き換える
        headers = {"Authorization": f"Bearer {access_token}"}
        # send massage
        text = f'{nothing_subject}\n{nothing_source}'
        data = {"message": f"\n{text}"}
        requests.post("https://notify-api.line.me/api/notify", data=data, headers=headers)

def email_send(nothing):
    nothing_number = len(nothing.index)

    for x in range(nothing_number):
        nothing_subtitle = nothing.iloc[x,4]
        nothing_subject = nothing.iloc[x,5]
        nothing_target = nothing.iloc[x,1]
        nothing_source = nothing.iloc[x,3]
        nothing_term = nothing.iloc[x,2]

        email_nothing_subject = nothing_subject.replace('\n','<br>')

        email_text = f'{email_nothing_subject}\n{nothing_source}'
        # SMTP認証情報
        account = "dezcz.python@gmail.com"
        password = "Nakasan33"
        # 送受信先
        to_email = "e2006422@soka-u.jp"
        from_email = "dezcz.python@gmail.com"

        # MIMEの作成
        subject = nothing_subtitle
        message = email_text
        msg = MIMEText(message, "html")
        msg["Subject"] = subject
        msg["To"] = to_email
        msg["From"] = from_email

        server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context())

        server.login(account, password)
        server.send_message(msg)



options = webdriver.ChromeOptions()
options.add_argument("--headless")

browser = webdriver.Chrome('/Users/nakamurayusuke/.conda/envs/pythonProject1/lib/python3.10/site-packages/chromedriver_binary/chromedriver', options=options)

#browser = webdriver.Chrome()

url = 'https://plas.soka.ac.jp/csp/plas/login.csp'
browser.get(url)

id = 'e2006422'
password = 'Soka7034'

elem_id = browser.find_element_by_name('un')
elem_pw = browser.find_element_by_name('pw')

elem_id.send_keys(id)
elem_pw.send_keys(password)

# ログインボタンを押す
time.sleep(1)
browser.find_element_by_name('Submit').click()
time.sleep(1)
browser.find_element_by_xpath('//*[@id="content"]/div/div[4]/div[2]/div/a').click()
time.sleep(1)

dfs_jugyou = scraping()

#学生生活・その他へ
tab_label = browser.find_elements_by_class_name('tab-label')
tab_label[1].click()
time.sleep(1)

dfs_seikatu = scraping()

#キャリア・教職・資格等へ
tab_label = browser.find_elements_by_class_name('tab-label')
tab_label[2].click()
time.sleep(1)

dfs_career = scraping()

time.sleep(2)
dfs = pd.concat([dfs_jugyou, dfs_seikatu, dfs_career], ignore_index= True)
dfs.index = dfs.index + 1

# 以下google spreadsheetにアクセス
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# 認証情報設定
# ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)

# OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

# 共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY = '1LjXf4j_xxGM6mj6vclrFKz0X4vlrMFRVmxO9yQMkhJk'

# 共有設定したスプレッドシートのシート1を開く
worksheet = gc.open_by_key(SPREADSHEET_KEY).sheet1

#google spreadsheetをdataframeに追加
df_logs = pd.DataFrame(worksheet.get_all_values())
df_logs.columns = ['タイトル', '対象', '掲示期間', '発信元', 'サブタイトル', '本文']
df_logs.drop(0, inplace=True)

nothing = dfs[~dfs['サブタイトル'].isin(df_logs['サブタイトル'])]

if not nothing.empty:
    send_line(nothing)
    email_send(nothing)

    num = len(df_logs.index) + 2

    col_lastnum = len(nothing.columns)  # DataFrameの列数
    row_lastnum = len(nothing.index)  # DataFrameの行数

    cell_list = worksheet.range('A' + str(num) + ':F' + str(row_lastnum + num - 1))
    for cell in cell_list:
        val = nothing.iloc[cell.row - num][cell.col - 1]
        cell.value = val

    worksheet.update_cells(cell_list)

else:
    pass

browser.close()