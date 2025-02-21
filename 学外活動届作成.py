import streamlit as st
from docx import Document
from datetime import datetime
import csv
import os

def write_csv_to_word_table(csv_file, word_file):
    """CSVデータをWordの指定された表のセルに書き込む（同一セルに連結、通し番号付き、件名も追記）"""
    try:
        document = Document(word_file)
    except FileNotFoundError:
        print(f"エラー: {word_file} が見つかりません。")
        return

    try:
        table = document.tables[2]  # 3つ目の表
    except IndexError:
        print("エラー: 3つ目の表が見つかりません。")
        return

    with open(csv_file, 'r', encoding='utf-8-sig') as infile:
        reader = csv.DictReader(infile)

        data_to_write = {  # 書き込むデータを辞書で管理
            (1, 1): "",  # 2行目2列目（件名）
            (2, 1): "",  # 3行目2列目（日時）
            (3, 2): "",  # 4行目3列目（場所名）
            (4, 2): "",  # 5行目3列目（場所住所）
        }
        item_number = 1

        for row in reader:
            try:
                event_date = datetime.strptime(row['開始日'], '%Y/%m/%d').strftime('%Y/%m/%d')
                event_time_start = row['開始時刻']
                event_time_end = row['終了時刻']
                event_place_name = row['場所'].split(',')[0] if row['場所'] else ""
                
                # 住所から「日本、」を削除
                address_parts = row['場所'].split(',')
                event_place_address = ','.join(address_parts[1:]).lstrip() if len(address_parts) > 1 else ""
                
                event_datetime = f"{event_date} {event_time_start}-{event_time_end}"
                event_subject = row['件名']

                # データを連結（通し番号付き）
                data_to_write[(1, 1)] += f"({item_number}) {event_subject}\n"
                data_to_write[(2, 1)] += f"({item_number}) {event_datetime}\n"
                data_to_write[(3, 2)] += f"({item_number}) {event_place_name}\n"
                data_to_write[(4, 2)] += f"({item_number}) {event_place_address}\n"

                item_number += 1

            except ValueError as e:
                print(f"日付変換エラー: {row} - {e}")
            except IndexError as e:
                print(f"表の範囲外への書き込み: {row} - {e}")
            except AttributeError as e:
                print(f"場所データの処理でエラーが発生しました。データを確認してください: {row} - {e}")
            except KeyError as e:
                print(f"CSVデータにキー {e} が見つかりません。")

        # まとめてセルに書き込み
        for (row_index, col_index), text in data_to_write.items():
            try:
                table.cell(row_index, col_index).text = text.rstrip('\n')
            except IndexError:
                print(f"表の範囲外への書き込みが発生しました。行:{row_index+1}, 列:{col_index+1} 表の行数を確認してください。")

    document.save(word_file)

st.title('Wordファイル書き込みツール')

uploaded_csv = st.file_uploader("CSVファイルをアップロードしてください", type="csv")
uploaded_word = st.file_uploader("Wordファイルをアップロードしてください", type="docx")

if uploaded_csv is not None and uploaded_word is not None:
    # ファイルを一時ファイルに保存
    with open("temp.csv", "wb") as f:
        f.write(uploaded_csv.getbuffer())
    with open("temp.docx", "wb") as f:
        f.write(uploaded_word.getbuffer())

    if st.button('書き込み'):
        with st.spinner('処理中...'): # スピナーを表示
            try:
                write_csv_to_word_table("temp.csv", "temp.docx")
                with open("temp.docx", "rb") as f:
                    st.download_button(label="ダウンロード", data=f, file_name="katsudoutodoke.docx") #ダウンロードボタン
                st.success("書き込みが完了しました！") # 成功メッセージを表示
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
            finally:
                os.remove("temp.csv")
                os.remove("temp.docx")
