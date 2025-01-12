import streamlit as st
import csv
from datetime import datetime
import os

def sort_csv_by_date(input_file, output_file):
    """CSVファイルを開始日順に並べ替える"""
    data = []
    try:
        with open(input_file, 'r', encoding='utf-8') as infile:
            reader = csv.DictReader(infile)
            for row in reader:
                try:
                    row['開始日'] = datetime.strptime(row['開始日'], '%Y/%m/%d')
                    data.append(row)
                except ValueError:
                    return "日付形式が正しくありません。'YYYY/MM/DD'の形式で入力してください。"
                except KeyError:
                    return "CSVファイルに'開始日'列が存在しません。"
    except FileNotFoundError:
        return "CSVファイルが見つかりません。"
    except Exception as e:
        return f"予期せぬエラーが発生しました: {e}"

    data.sort(key=lambda x: x['開始日'])

    with open(output_file, 'w', encoding='utf-8', newline='') as outfile:
        fieldnames = reader.fieldnames
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()
        for row in data:
            row['開始日'] = row['開始日'].strftime('%Y/%m/%d')
            writer.writerow(row)
    return None

st.title('CSVソートツール')

uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type="csv")

if uploaded_file is not None:
    with open("temp.csv", "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button('ソート'):
        with st.spinner('処理中...'):
            error_message = sort_csv_by_date("temp.csv", "sorted.csv")
            if error_message:
                st.error(error_message)
            else:
                with open("sorted.csv", "rb") as f:
                    st.download_button(label="ダウンロード", data=f, file_name="sorted.csv")
                st.success("ソートが完了しました！")
        os.remove("temp.csv")
        os.remove("sorted.csv")