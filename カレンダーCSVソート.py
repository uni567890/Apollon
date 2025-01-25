import streamlit as st
import csv
from datetime import datetime
import os

def sort_csv_by_date_and_time(input_file, output_file):
    """
    CSVファイルを開始日と開始時刻順に並べ替える

    Args:
        input_file (str): 入力CSVファイルのパス
        output_file (str): 出力CSVファイルのパス
    """
    data = []
    with open(input_file, 'r', encoding='utf-8-sig') as infile:
        reader = csv.DictReader(infile)
        for row in reader:
            try:
                row['開始日'] = datetime.strptime(row['開始日'], '%Y/%m/%d').date() # 日付型に変換
                row['開始時刻'] = datetime.strptime(row['開始時刻'], '%H:%M:%S').time() # 時刻型に変換
                data.append(row)
            except ValueError as e:
                print(f"警告: 不正な日付または時刻形式の行があります: {row} - {e}")

    # 開始日と開始時刻でソート
    data.sort(key=lambda x: (x['開始日'], x['開始時刻']))

    with open(output_file, 'w', encoding='utf-8-sig', newline='') as outfile:
        fieldnames = reader.fieldnames
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            # 時刻のフォーマットを変更（秒を削除）
            writer.writerow({
                '件名': row['件名'],
                '開始日': row['開始日'].strftime('%Y/%m/%d'),
                '開始時刻': row['開始時刻'].strftime('%H:%M'), # 秒を削除
                '終了時刻': row['終了時刻'].split(':')[0] + ':' + row['終了時刻'].split(':')[1] if row.get('終了時刻') else '', # 秒を削除, Noneチェック追加
                '場所': row['場所']
            })

st.title('CSVソートツール')

uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type="csv")

if uploaded_file is not None:
    with open("temp.csv", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button('ソート'):
        with st.spinner('処理中...'):
            error_message = sort_csv_by_date_and_time("temp.csv", "sorted.csv")
            if error_message:
                st.error(error_message)
            else:
                with open("sorted.csv", "rb") as f:
                    st.download_button(label="ダウンロード", data=f, file_name="sorted.csv")
                st.success("ソートが完了しました！")
        os.remove("temp.csv")
        os.remove("sorted.csv")
