from flask import Flask, render_template, request, send_file
import pandas as pd
import os
import re
from datetime import datetime

app = Flask(__name__)
# app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024  # 限制上傳大小 2MB

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ========== EXP 檔案處理邏輯 ==========
def process_exp_file(file_storage):
    lines = [line.strip() for line in file_storage.stream.read().decode("utf-8", errors="ignore").splitlines() if line.strip()]
    headers = [h.strip('"') for h in lines[1].split('\t')]
    data_lines = lines[2:]
    data = [[cell.strip('"') for cell in row.split('\t')] for row in data_lines]

    max_cols = len(headers)
    data = [row[:max_cols] + [''] * (max_cols - len(row)) for row in data]
    df = pd.DataFrame(data, columns=headers)

    def clean_illegal_chars(val):
        if isinstance(val, str):
            return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]', '', val)
        return val

    df = df.applymap(clean_illegal_chars)

    wanted_columns = ["ID", "Part Reference", "Value", "Net Name"]

    def normalize(col_name):
        return col_name.strip().lower().replace(" ", "").replace("_", "")

    normalized_map = {normalize(col): col for col in df.columns}
    final_cols = [normalized_map[nc] for nc in map(normalize, wanted_columns) if nc in normalized_map]
    df = df[final_cols]

    df['Part Reference'] = df['Part Reference'].replace({'<null>': '', None: ''})
    df['Value'] = df['Value'].replace({'<null>': '', None: ''})

    df['main_id'] = df['ID'].str.split(':').str[0]
    main_df = df[~df['ID'].str.contains(":")].set_index('ID')
    child_df = df[df['ID'].str.contains(":")].copy()

    child_df = child_df.merge(main_df[['Part Reference', 'Value']], how='left', left_on='main_id', right_index=True, suffixes=('', '_main'))
    child_df['Part Reference'] = child_df['Part Reference'].replace('', child_df['Part Reference_main'])
    child_df['Value'] = child_df['Value_main']

    return child_df.drop(columns=['main_id', 'Part Reference_main', 'Value_main'])

# ========== 前端頁面 ==========
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file1 = request.files.get("file1")
        file2 = request.files.get("file2")
        primary_key = request.form.get("primary_key")
        ignore_case = request.form.get("ignore_case") == "on"

        if not file1 or not file2 or not primary_key:
            return "❌ 請上傳兩個 .EXP 檔案並選擇主鍵", 400

        try:
            df1 = process_exp_file(file1)
            df2 = process_exp_file(file2)

            if ignore_case:
                df1 = df1.applymap(lambda x: str(x).strip().upper())
                df2 = df2.applymap(lambda x: str(x).strip().upper())

            df1.set_index(primary_key, inplace=True)
            df2.set_index(primary_key, inplace=True)
            df1 = df1[~df1.index.duplicated(keep='first')]
            df2 = df2[~df2.index.duplicated(keep='first')]

            added_keys = df2.index.difference(df1.index)
            removed_keys = df1.index.difference(df2.index)
            common_keys = df1.index.intersection(df2.index)

            df_added = df2.loc[added_keys].copy()
            df_removed = df1.loc[removed_keys].copy()
            df_modified = []

            for key in common_keys:
                old_row = df1.loc[key]
                new_row = df2.loc[key]
                diff = old_row != new_row
                if diff.any():
                    combined = new_row.copy()
                    for col in df1.columns:
                        if diff[col]:
                            combined[col] = f"{old_row[col]} → {new_row[col]}"
                    df_modified.append(combined)

            df_modified = pd.DataFrame(df_modified)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(UPLOAD_FOLDER, f"Compare_Result_{ts}.xlsx")

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                df_added.reset_index().to_excel(writer, sheet_name="新增列", index=False)
                df_removed.reset_index().to_excel(writer, sheet_name="刪除列", index=False)
                df_modified.reset_index().to_excel(writer, sheet_name="修改列", index=False)

            return send_file(output_path, as_attachment=True)

        except Exception as e:
            return f"❌ 比對過程中發生錯誤：{e}", 500

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
