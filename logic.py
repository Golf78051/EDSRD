# logic.py
import pandas as pd
import os
import re
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook


def clean_illegal_chars(val):
    if isinstance(val, str):
        return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F]', '', val)
    return val


def process_exp_and_trace(exp_file_path, max_depth, source_input, output_folder):
    output_file = os.path.join(output_folder, "Exp_Processed.xlsx")
    summary_path = os.path.join(output_folder, "No power source net.xlsx")
    trace_path = os.path.join(output_folder, "Power_Trace_Detail.xlsx")

    # ========== Part 1: EXP to Excel ==========
    with open(exp_file_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = [line.strip() for line in f if line.strip()]

    headers = [h.strip('"') for h in lines[1].split('\t')]
    data_lines = lines[2:]
    data = [[cell.strip('"') for cell in row.split('\t')] for row in data_lines]

    max_cols = len(headers)
    data = [row[:max_cols] + [''] * (max_cols - len(row)) for row in data]
    df = pd.DataFrame(data, columns=headers)
    df = df.applymap(clean_illegal_chars)

    wanted_columns = ["ID", "Part Reference", "Value", "Net Name"]
    def normalize(col_name):
        return col_name.strip().lower().replace(" ", "").replace("_", "")
    existing_cols = df.columns
    normalized_map = {normalize(col): col for col in existing_cols}
    final_cols = [normalized_map[nc] for nc in map(normalize, wanted_columns) if nc in normalized_map]
    df = df[final_cols]

    df['Part Reference'] = df['Part Reference'].replace({'<null>': '', None: ''})
    df['Value'] = df['Value'].replace({'<null>': '', None: ''})

    main_part_ref_map, main_value_map, new_rows = {}, {}, []
    for idx, row in df.iterrows():
        id_value = str(row['ID']).strip()
        if ':' not in id_value:
            main_part_ref_map[id_value] = row['Part Reference']
            main_value_map[id_value] = row['Value']
        else:
            main_id = id_value.split(':')[0]
            if row['Part Reference'] == '':
                row['Part Reference'] = main_part_ref_map.get(main_id, '')
            row['Value'] = main_value_map.get(main_id, row['Value'])
            new_rows.append(row)
    new_df = pd.DataFrame(new_rows)
    new_df.to_excel(output_file, index=False)

    # ========== Part 2: 呼叫外部追蹤邏輯（由原程式邏輯整合而來） ==========
    import subprocess
    import sys
    subprocess.run([sys.executable, os.path.join(os.path.dirname(__file__), "power_trace_core.py"), 
                    output_file, str(max_depth), source_input, summary_path, trace_path])

    return summary_path, trace_path
