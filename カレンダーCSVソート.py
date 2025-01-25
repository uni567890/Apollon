import streamlit as st
import csv
from datetime import datetime
import os

def sort_and_filter_csv(input_file, output_file, selected_month):
    """
    CSVファイルを開始日と開始時刻順に並べ替え、指定した月のデータのみを抽出し、不要なカラムを削除する。

    Args:
        input_file (str): 入力CSVファイルのパス
        output_file (str): 出力CSVファイルのパス
        selected_month (datetime.date): 抽出する月
    """
    data = []
    with open(input_file, 'r', encoding='utf-8-sig') as infile:
        reader = csv.DictReader(infile)
        for row in reader:
            try:
                row['開始日'] = datetime.strptime(row['開始日'], '%Y/%m/%d').date()
                row['開始時刻'] = datetime.strptime(row['開始時刻'], '%H:%M:%S').time()
                row['終了時刻'] = datetime.strptime(row['終了時刻'], '%H:%M:%S').time()

                # 月でフィルタ
                if row['開始日'].year == selected_month.year and row['開始日'].month == selected_month.month:
                    data.append(row)
            except ValueError as e:
                print(f"警告: 不正な日付または時刻形式の行があります: {row} - {e}")
            except KeyError as e:
                print(f"警告: 必要なカラムが存在しません: {e}")

    # 開始日と開始時刻でソート
    data.sort(key=lambda x: (x['開始日'], x['開始時刻']))

    with open(output_file, 'w', encoding='utf-8-sig', newline='') as outfile:
        fieldnames = ['件名', '開始日', '開始時刻', '終了時刻', '場所']  # 出力するカラムを指定
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            writer.writerow({
                '件名': row['件名'],
                '開始日': row['開始日'].strftime('%Y/%m/%d'),
                '開始時刻': row['開始時刻'].strftime('%H:%M'), # 秒を削除
                '終了時刻': row['終了時刻'].strftime('%H:%M'),  # 秒を削除
                '場所': row['場所']
            })

st.title('CSVソート＆フィルタツール')

uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type="csv")
month = st.date_input("抽出する月を選択", value=datetime.now()) # 月選択ウィジェットを追加

if uploaded_file is not None:
    with open("temp.csv", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button('ソート＆フィルタ'):
        with st.spinner('処理中...'):
            sort_and_filter_csv("temp.csv", "filtered_sorted.csv", month)
            with open("filtered_sorted.csv", "rb") as f:
                st.download_button(label="ダウンロード", data=f, file_name="filtered_sorted.csv")
            st.success("ソートとフィルタが完了しました！")
        os.remove("temp.csv")
        os.remove("filtered_sorted.csv")
