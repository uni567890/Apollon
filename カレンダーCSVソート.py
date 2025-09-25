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
        # ヘッダーを読んで形式を判別
        header = infile.readline().strip().split(',')
        infile.seek(0)  # ファイルポインタを先頭に戻す
        reader = csv.DictReader(infile)

        # 新しい形式のヘッダーかチェック
        is_new_format = 'StartDate' in header and 'Summary' in header

        for row in reader:
            try:
                if is_new_format:
                    if not row.get('StartDate'):
                        continue
                    start_datetime = datetime.strptime(row['StartDate'], '%Y/%m/%d %H:%M')
                    end_datetime_str = row.get('EndDate', '')
                    end_time = datetime.strptime(end_datetime_str, '%Y/%m/%d %H:%M').time() if end_datetime_str else ''

                    processed_row = {
                        '件名': row.get('Summary', ''),
                        '開始日': start_datetime.date(),
                        '開始時刻': start_datetime.time(),
                        '終了時刻': end_time,
                        '場所': row.get('Location', '')
                    }
                else: # 既存の形式
                    if not row.get('開始日') or not row.get('開始時刻'):
                        continue
                    
                    end_time_str = row.get('終了時刻', '')
                    end_time = datetime.strptime(end_time_str, '%H:%M:%S').time() if end_time_str else ''

                    processed_row = {
                        '件名': row.get('件名', ''),
                        '開始日': datetime.strptime(row['開始日'], '%Y/%m/%d').date(),
                        '開始時刻': datetime.strptime(row['開始時刻'], '%H:%M:%S').time(),
                        '終了時刻': end_time,
                        '場所': row.get('場所', '')
                    }
                data.append(processed_row)
            except (ValueError, KeyError) as e:
                print(f"警告: スキップされた行があります: {row} - {e}")


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
                    '終了時刻': row['終了時刻'].strftime('%H:%M') if row.get('終了時刻') else '',
                    '場所': row['場所']
                })

st.title('CSVソートツール')

uploaded_file = st.file_uploader("CSVファイルをアップロードしてください", type="csv")

if uploaded_file is not None:
    # BOMが含まれている場合でも読み込めるように utf-8-sig を指定
    file_content = uploaded_file.getvalue().decode('utf-8-sig')
    with open("temp.csv", "w", encoding="utf-8-sig") as f:
        f.write(file_content)

    # 月を選択するselectboxを追加
    months = list(range(1, 13))
    # 現在の月をデフォルトに設定
    default_month = datetime.now().month
    selected_month = st.selectbox("出力する月を選択してください", months, index=months.index(default_month))

    if st.button('ソート'):
        with st.spinner('処理中...'):
            sort_csv_by_date_and_time("temp.csv", "sorted.csv", selected_month)
            with open("sorted.csv", "rb") as f:
                st.download_button(label="ダウンロード", data=f, file_name=f"{selected_month}月度練習日程.csv", mime="text/csv")
            st.success("ソートが完了しました！")
        try:
            os.remove("temp.csv")
            os.remove("sorted.csv")
        except FileNotFoundError:
            pass
