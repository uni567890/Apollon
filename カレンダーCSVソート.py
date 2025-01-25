import streamlit as st
import csv
from datetime import datetime
import os

def sort_csv_by_date_and_time(input_file, output_file, selected_month):
    """
    CSVファイルを開始日と開始時刻順に並べ替え、指定カラム以外を削除し、指定月のデータのみ出力

    Args:
        input_file (str): 入力CSVファイルのパス
        output_file (str): 出力CSVファイルのパス
        selected_month (int): 出力する月 (1-12)
    """
    data = []
    with open(input_file, 'r', encoding='utf-8-sig') as infile:
        reader = csv.DictReader(infile)
        for row in reader:
            try:
                row['開始日'] = datetime.strptime(row['開始日'], '%Y/%m/%d').date()
                row['開始時刻'] = datetime.strptime(row['開始時刻'], '%H:%M:%S').time()
                data.append(row)
            except ValueError as e:
                print(f"警告: 不正な日付または時刻形式の行があります: {row} - {e}")

    data.sort(key=lambda x: (x['開始日'], x['開始時刻']))

    with open(output_file, 'w', encoding='utf-8-sig', newline='') as outfile:
        fieldnames = ['件名', '開始日', '開始時刻', '終了時刻', '場所']
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
          # 指定された月と一致するかチェック
            if row['開始日'].month == selected_month:
                writer.writerow({
                    '件名': row['件名'],
                    '開始日': row['開始日'].strftime('%Y/%m/%d'),
                    '開始時刻': row['開始時刻'].strftime('%H:%M'),
                    '終了時刻': row['終了時刻'].split(':')[0] + ':' + row['終了時刻'].split(':')[1] if row.get('終了時刻') else '',
                    '場所': row['場所']
                })

st.title('CSVソートツール')

uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type="csv")

if uploaded_file is not None:
    with open("temp.csv", "wb") as f:
        f.write(uploaded_file.getbuffer())

    # 月を選択するselectboxを追加
    months = list(range(1, 13))
    selected_month = st.selectbox("出力する月を選択してください", months)

    if st.button('ソート'):
        with st.spinner('処理中...'):
            error_message = sort_csv_by_date_and_time("temp.csv", "sorted.csv", selected_month)
            if error_message:
                st.error(error_message)
            else:
                with open("sorted.csv", "rb") as f:
                    st.download_button(label="ダウンロード", data=f, file_name="sorted.csv")
                st.success("ソートが完了しました！")
        try:
            os.remove("temp.csv")
            os.remove("sorted.csv")
        except FileNotFoundError:
            pass
